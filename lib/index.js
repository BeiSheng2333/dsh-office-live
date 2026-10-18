/**
 * dsh-office-live
 *
 * Model-facing tools that drive the Microsoft Office applications the user is
 * looking at *right now*: read the live document contents, capture a screenshot
 * of the application window (so the model can "see" the layout), and write
 * changes back into that same window without touching files.
 *
 * The heavy lifting lives in `assets/office-live.ps1` (Windows PowerShell 5.1+),
 * which talks to Office through COM and grabs window pixels through user32 +
 * System.Drawing. This module only marshals JSON across that boundary.
 */
import { execFile } from 'node:child_process';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { defineTool } from '@deepseek-ai/dsh-tools';

const name = 'dsh-office-live';
const inject = ['tools'];

const SCRIPT_PATH = fileURLToPath(new URL('../assets/office-live.ps1', import.meta.url));
const TARGETS = ['word', 'excel', 'powerpoint'];
const WORK_DIR = join(tmpdir(), 'dsh-office-live');
const BRIDGE_TIMEOUT_MS = 180_000;

const OUTPUT_SCHEMA = { type: 'object', additionalProperties: true };

const targetParameter = {
  type: 'string',
  enum: TARGETS,
  description: 'Which running Office application to act on. Defaults to "word".',
};

let serial = 0;
function tempPath(prefix, extension) {
  serial += 1;
  return join(WORK_DIR, `${prefix}-${Date.now()}-${serial}${extension}`);
}

function powershellExecutable() {
  return process.env.OFFICE_LIVE_POWERSHELL || 'powershell.exe';
}

function parseBridgeOutput(text) {
  const trimmed = text.replace(/^\uFEFF/, '').trim();
  if (trimmed === '') throw new Error('office-live bridge produced no output');
  let value;
  try {
    value = JSON.parse(trimmed);
  } catch {
    throw new Error(`office-live bridge returned malformed JSON: ${trimmed.slice(0, 800)}`);
  }
  if (value !== null && typeof value === 'object' && value.error) {
    const error = value.error;
    const code = typeof error.code === 'string' ? error.code : 'error';
    const message = typeof error.message === 'string' ? error.message : JSON.stringify(error);
    throw new Error(`office-live ${code}: ${message}`);
  }
  return value;
}

/**
 * Run one bridge action and return its parsed JSON result.
 * The result is written to a UTF-8 side file so Chinese document text survives
 * the Windows console code page, then the file is removed.
 */
async function runBridge(args, exec) {
  await mkdir(WORK_DIR, { recursive: true });
  const resultFile = tempPath('result', '.json');
  try {
    await new Promise((resolve, reject) => {
      const child = execFile(
        powershellExecutable(),
        [
          '-NoProfile',
          '-NonInteractive',
          '-ExecutionPolicy',
          'Bypass',
          '-File',
          SCRIPT_PATH,
          ...args,
          '-ResultFile',
          resultFile,
        ],
        { windowsHide: true, maxBuffer: 16 * 1024 * 1024, timeout: BRIDGE_TIMEOUT_MS },
        (error, _stdout, stderr) => {
          if (error) {
            const detail = stderr ? ` | ${String(stderr).trim().slice(0, 2000)}` : '';
            reject(new Error(`office-live bridge failed: ${error.message}${detail}`));
            return;
          }
          resolve(undefined);
        },
      );
      exec?.signal?.addEventListener?.('abort', () => child.kill(), { once: true });
    });
    return parseBridgeOutput(await readFile(resultFile, 'utf8'));
  } finally {
    await rm(resultFile, { force: true });
  }
}

function textResult(value) {
  return [{ type: 'text', text: JSON.stringify(value, null, 2) }];
}

function registerStatus(ctx) {
  return ctx.tools.register(
    defineTool({
      name: 'office_live_status',
      description:
        'List whether Word, Excel and PowerPoint are currently running in the user\'s desktop session, and which documents/workbooks/presentations are open. Call this first when the user refers to "the document I have open" so you know which application and file they mean.',
      parameters: {},
      output: { schema: OUTPUT_SCHEMA, render: (_args, value) => textResult(value) },
      async execute(_args, exec) {
        return runBridge(['-Action', 'status'], exec);
      },
    }),
  );
}

function registerRead(ctx) {
  return ctx.tools.register(
    defineTool({
      name: 'office_live_read',
      description:
        'Read the LIVE contents of the document currently open in Word, Excel or PowerPoint (attaches to the running application, not to a file on disk). Word returns paragraphs with style/bold/font/size plus tables, page and word counts and the current selection (first 400 paragraphs). Excel returns sheets and the active sheet\'s used range as a cell matrix with evaluated values (first 200 rows x 40 columns). PowerPoint returns the active slide and each shape with position and text. Changes the user makes in the window are visible to the next read.',
      parameters: {
        target: targetParameter,
      },
      output: { schema: OUTPUT_SCHEMA, render: (_args, value) => textResult(value) },
      async execute(args, exec) {
        return runBridge(['-Action', 'read', '-Target', args.target ?? 'word'], exec);
      },
    }),
  );
}

function registerShot(ctx) {
  return ctx.tools.register(
    defineTool({
      name: 'office_live_shot',
      description:
        'Capture a PNG screenshot of the running Office application window (ribbon, page and all) and return the image path. Use it to see the current layout/formatting before editing, and again afterwards to verify the change. The window is brought to the foreground first. Read the returned path with read_image.',
      parameters: {
        target: targetParameter,
        out: {
          type: 'string',
          description: 'Optional PNG path to write. Defaults to a file under the system temp directory.',
        },
      },
      output: { schema: OUTPUT_SCHEMA, render: (_args, value) => textResult(value) },
      async execute(args, exec) {
        const out = args.out ?? tempPath('shot', '.png');
        return runBridge(['-Action', 'shot', '-Target', args.target ?? 'word', '-Out', out], exec);
      },
    }),
  );
}

function registerSnapshot(ctx) {
  return ctx.tools.register(
    defineTool({
      name: 'office_live_snapshot',
      description:
        'One call that both reads the live document structure and captures a window screenshot (read + shot). Prefer this before editing so you have the exact text and the visual layout at the same time.',
      parameters: {
        target: targetParameter,
        out: {
          type: 'string',
          description: 'Optional PNG path to write. Defaults to a file under the system temp directory.',
        },
      },
      output: { schema: OUTPUT_SCHEMA, render: (_args, value) => textResult(value) },
      async execute(args, exec) {
        const out = args.out ?? tempPath('snap', '.png');
        return runBridge(['-Action', 'snapshot', '-Target', args.target ?? 'word', '-Out', out], exec);
      },
    }),
  );
}

function registerEdit(ctx) {
  return ctx.tools.register(
    defineTool({
      name: 'office_live_edit',
      description:
        'Apply a batch of edits to the document the user has open, live in the running Office window. Nothing is saved to disk unless save is true. Ops: word.append_paragraph {text, style?, bold?, size?, font?, align?}; word.insert_paragraph_before {index, text, style?}; word.set_paragraph {index, text?, style?, bold?, size?, font?, align?, color?}; word.replace {find, replace, all?}; word.insert_table {data:[[...]], header?}; word.insert_page_break; excel.set_cell {cell, value?|formula?, bold?, numberFormat?, sheet?}; excel.set_range {startCell, data, sheet?}; excel.add_sheet {name}; ppt.set_text {slide?, shape?, text, size?, bold?}; ppt.add_slide {layout?, title?}. Each op reports ok; a failed op does not abort the rest.',
      parameters: {
        target: targetParameter,
        ops: {
          type: 'array',
          required: true,
          items: { type: 'object' },
          description: 'Ordered edit operations, each an object with an "op" field and its arguments.',
        },
        save: {
          type: 'boolean',
          description: 'Save the document after all ops succeed. Defaults to false (changes stay unsaved in the window).',
        },
      },
      output: { schema: OUTPUT_SCHEMA, render: (_args, value) => textResult(value) },
      async execute(args, exec) {
        const opsFile = tempPath('ops', '.json');
        await writeFile(opsFile, JSON.stringify(args.ops), 'utf8');
        try {
          const argv = ['-Action', 'edit', '-Target', args.target ?? 'word', '-Ops', opsFile];
          if (args.save === true) argv.push('-Save');
          return await runBridge(argv, exec);
        } finally {
          await rm(opsFile, { force: true });
        }
      },
    }),
  );
}

function registerOpen(ctx) {
  return ctx.tools.register(
    defineTool({
      name: 'office_live_open',
      description:
        'Open a file in the real Word / Excel / PowerPoint application so later office_live_* calls can read and edit it live. Prefer this over file-level tools when the user wants to watch the editing happen on screen.',
      parameters: {
        target: targetParameter,
        file: {
          type: 'string',
          required: true,
          description: 'Absolute or workspace-relative path of the .docx / .xlsx / .pptx to open.',
        },
      },
      output: { schema: OUTPUT_SCHEMA, render: (_args, value) => textResult(value) },
      async execute(args, exec) {
        return runBridge(['-Action', 'open', '-Target', args.target ?? 'word', '-File', args.file], exec);
      },
    }),
  );
}

function registerClose(ctx) {
  return ctx.tools.register(
    defineTool({
      name: 'office_live_close',
      description:
        'Close the document currently open in the Office window. Without save the window closes and unsaved changes are discarded; pass save to save first.',
      parameters: {
        target: targetParameter,
        save: {
          type: 'boolean',
          description: 'Save before closing. Defaults to false.',
        },
      },
      output: { schema: OUTPUT_SCHEMA, render: (_args, value) => textResult(value) },
      async execute(args, exec) {
        const argv = ['-Action', 'close', '-Target', args.target ?? 'word'];
        if (args.save === true) argv.push('-Save');
        return runBridge(argv, exec);
      },
    }),
  );
}

function apply(ctx) {
  ctx.effect(() => {
    const disposers = [
      registerStatus(ctx),
      registerRead(ctx),
      registerShot(ctx),
      registerSnapshot(ctx),
      registerEdit(ctx),
      registerOpen(ctx),
      registerClose(ctx),
    ];
    return () => {
      for (const dispose of disposers) dispose();
    };
  });
}

export { apply, inject, name };
