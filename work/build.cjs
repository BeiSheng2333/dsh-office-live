const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak,
  HeightRule, TabStopType, LineRuleType, TableLayoutType
} = require("docx");

const CJK = "Microsoft YaHei";
const FONT = { ascii: "Arial", hAnsi: "Arial", eastAsia: CJK };
const NAVY = "1F4E79";
const BLUE = "2E75B6";
const GRAY = "595959";
const LIGHT = "F2F7FB";
const CONTENT_W = 9026;

const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: "C9D6E2" };
const allThin = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };

const noBorder = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function text(t, o = {}) {
  return new TextRun({ text: t, font: FONT, size: 22, ...o });
}

function body(t, o = {}) {
  return new Paragraph({
    spacing: { after: 140, line: 336, lineRule: LineRuleType.AUTO },
    children: [text(t, o)],
  });
}

function h1(t, newPage = false) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    pageBreakBefore: newPage,
    keepNext: true,
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 6 } },
    children: [new TextRun({ text: t, font: FONT, size: 32, bold: true, color: NAVY })],
  });
}

function h2(t, newPage = false) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    pageBreakBefore: newPage,
    keepNext: true,
    children: [new TextRun({ text: t, font: FONT, size: 26, bold: true, color: BLUE })],
  });
}

function spacer(before = 120, after = 0) {
  return new Paragraph({ spacing: { before, after }, children: [] });
}

function step(ref, lead, desc) {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 100, line: 300, lineRule: LineRuleType.AUTO },
    children: [
      text(lead, { bold: true, color: "24425F" }),
      text(desc),
    ],
  });
}

function caption(t) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 260 },
    children: [new TextRun({ text: t, font: FONT, size: 18, color: GRAY })],
  });
}

function photoSlot(figCaption, slotW = 4400, slotH = 3000, hint = "请替换为现场实拍照片（建议比例 4:3）") {
  const dashed = { style: BorderStyle.DASHED, size: 6, color: "AEB9C4" };
  return [
    new Table({
      width: { size: slotW, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: [slotW],
      alignment: AlignmentType.CENTER,
      rows: [
        new TableRow({
          height: { value: slotH, rule: HeightRule.ATLEAST },
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: slotW, type: WidthType.DXA },
              borders: { top: dashed, bottom: dashed, left: dashed, right: dashed },
              shading: { fill: "F8F9FA", type: ShadingType.CLEAR },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 120, bottom: 120, left: 160, right: 160 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 60 },
                  children: [new TextRun({ text: "［ 照片预留位 ］", font: FONT, size: 20, bold: true, color: "93A1AF" })],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun({ text: hint, font: FONT, size: 16, color: "9AA6B2" })],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    caption(figCaption),
  ];
}

function photoPair(capL, capR) {
  const dashed = { style: BorderStyle.DASHED, size: 6, color: "AEB9C4" };
  const W = 4000, G = 1026;
  const phCell = () =>
    new TableCell({
      width: { size: W, type: WidthType.DXA },
      borders: { top: dashed, bottom: dashed, left: dashed, right: dashed },
      shading: { fill: "F8F9FA", type: ShadingType.CLEAR },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 120, bottom: 120, left: 160, right: 160 },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [new TextRun({ text: "［ 照片预留位 ］", font: FONT, size: 20, bold: true, color: "93A1AF" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "请替换为现场实拍照片（建议比例 4:3）", font: FONT, size: 16, color: "9AA6B2" })],
        }),
      ],
    });
  const gapCell = (rowSpan = false) =>
    new TableCell({
      width: { size: G, type: WidthType.DXA },
      borders: noBorder,
      children: [new Paragraph({ spacing: { after: 0 }, children: [] })],
    });
  const capCell = (t) =>
    new TableCell({
      width: { size: W, type: WidthType.DXA },
      borders: noBorder,
      margins: { top: 60, bottom: 200, left: 80, right: 80 },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: t, font: FONT, size: 18, color: GRAY })],
        }),
      ],
    });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    columnWidths: [W, G, W],
    alignment: AlignmentType.CENTER,
    rows: [
      new TableRow({
        cantSplit: true,
        height: { value: 3000, rule: HeightRule.ATLEAST },
        children: [phCell(), gapCell(), phCell()],
      }),
      new TableRow({
        cantSplit: true,
        children: [capCell(capL), gapCell(), capCell(capR)],
      }),
    ],
  });
}

function callout(title, items) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    columnWidths: [CONTENT_W],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: CONTENT_W, type: WidthType.DXA },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 4, color: "D5E3F0" },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5E3F0" },
              right: { style: BorderStyle.SINGLE, size: 4, color: "D5E3F0" },
              left: { style: BorderStyle.SINGLE, size: 18, color: NAVY },
            },
            shading: { fill: LIGHT, type: ShadingType.CLEAR },
            margins: { top: 140, bottom: 140, left: 220, right: 200 },
            children: [
              new Paragraph({
                spacing: { after: 80 },
                children: [new TextRun({ text: title, font: FONT, size: 20, bold: true, color: NAVY })],
              }),
              ...items.map((t) =>
                new Paragraph({
                  numbering: { reference: "bullets", level: 0 },
                  spacing: { after: 40, line: 290, lineRule: LineRuleType.AUTO },
                  children: [new TextRun({ text: t, font: FONT, size: 19, color: "33475B" })],
                })
              ),
            ],
          }),
        ],
      }),
    ],
  });
}

function headCell(t, w) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: allThin,
    shading: { fill: NAVY, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 100, bottom: 100, left: 130, right: 130 },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: t, font: FONT, size: 19, bold: true, color: "FFFFFF" })],
      }),
    ],
  });
}

function dataCell(t, w, fill, center = false, bold = false) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: allThin,
    shading: { fill, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 100, bottom: 100, left: 130, right: 130 },
    children: [
      new Paragraph({
        alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
        spacing: { line: 300, lineRule: LineRuleType.AUTO },
        children: [new TextRun({ text: t, font: FONT, size: 19, bold, color: "33475B" })],
      }),
    ],
  });
}

const numberingConfigs = [
  {
    reference: "bullets",
    levels: [
      {
        level: 0,
        format: LevelFormat.BULLET,
        text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 300, hanging: 200 } } },
      },
    ],
  },
  ...["pick", "box", "short", "move", "putaway", "std", "count", "abn"].map((r) => ({
    reference: r,
    levels: [
      {
        level: 0,
        format: LevelFormat.DECIMAL,
        text: "%1.",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 460, hanging: 340 } } },
      },
    ],
  })),
];

const infoRows = [
  ["适用区域", "一层横梁货架区（非设备区共用同一拣货功能）"],
  ["适用设备", "手持终端（PDA）"],
  ["涉及模块", "出库　/　库存　/　入库"],
  ["文档版本", "V1.0　（2026 年 11 月）"],
];

const overviewRows = [
  ["拣货作业", "出库 → 平置横梁拣货", "申领任务、扫码拣货、绑定托盘", "生成出库绑箱任务"],
  ["移库作业", "库存 → 分段移库 / 标准移库", "下架、上架确认或库位间转移", "库位与容器信息更新"],
  ["盘点作业", "库存 → 非设备区盘点", "扫起始码、系统推荐库位、实盘录入", "盘点结果确认"],
];

const rulesRows = [
  ["1", "申领限量", "单次最多申领 15 个拣货任务，申领后及时执行，避免任务积压"],
  ["2", "三重核对", "每步操作核对书号、数量、容器号，与实物一致后再点击确认"],
  ["3", "扫码优先", "库位、容器、书号均以扫码识别，减少手工录入造成的错漏"],
  ["4", "作业衔接", "拣货缺货后转入异常盘点；移库须完成上架确认后闭环"],
  ["5", "异常留痕", "缺货、盘点差异均以系统单据留痕，不在线下另行处理"],
  ["6", "照片留存", "各章节照片位用于补充现场实拍，便于培训与复核"],
];

const doc = new Document({
  title: "横梁货架区手持作业指导书",
  subject: "拣货、移库、盘点标准操作流程",
  description: "横梁货架区手持终端作业标准作业指导书（V1.0）",
  styles: {
    default: {
      document: {
        run: { font: FONT, size: 22, color: "262626" },
        paragraph: { spacing: { line: 300, lineRule: LineRuleType.AUTO } },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, color: NAVY, font: FONT },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0, keepNext: false, keepLines: false },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: BLUE, font: FONT },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1, keepNext: false, keepLines: false },
      },
    ],
  },
  numbering: { config: numberingConfigs },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "C9D6E2", space: 4 } },
              tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
              children: [
                new TextRun({ text: "横梁货架区手持作业指导书", font: FONT, size: 16, color: "7F8C9A" }),
                new TextRun({ text: "\tV1.0", font: FONT, size: 16, color: "7F8C9A" }),
              ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "第 ", font: FONT, size: 16, color: "7F8C9A" }),
                new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "7F8C9A" }),
                new TextRun({ text: " 页", font: FONT, size: 16, color: "7F8C9A" }),
              ],
            }),
          ],
        }),
      },
      children: [
        // Cover
        new Table({
          width: { size: CONTENT_W, type: WidthType.DXA },
          columnWidths: [CONTENT_W],
          rows: [
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: CONTENT_W, type: WidthType.DXA },
                  borders: noBorder,
                  shading: { fill: NAVY, type: ShadingType.CLEAR },
                  margins: { top: 480, bottom: 480, left: 420, right: 420 },
                  children: [
                    new Paragraph({
                      spacing: { after: 60 },
                      children: [new TextRun({ text: "手持终端作业规范", font: FONT, size: 18, color: "BFD7EE" })],
                    }),
                    new Paragraph({
                      spacing: { after: 100 },
                      children: [new TextRun({ text: "横梁货架区手持作业指导书", font: FONT, size: 42, bold: true, color: "FFFFFF" })],
                    }),
                    new Paragraph({
                      children: [new TextRun({ text: "拣货　·　移库　·　盘点　标准操作流程", font: FONT, size: 22, color: "D6E4F0" })],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
        spacer(360, 240),
        new Table({
          width: { size: CONTENT_W, type: WidthType.DXA },
          layout: TableLayoutType.FIXED,
          columnWidths: [1900, CONTENT_W - 1900],
          rows: infoRows.map(([k, v], i) =>
            new TableRow({
              cantSplit: true,
              children: [
                new TableCell({
                  width: { size: 1900, type: WidthType.DXA },
                  borders: allThin,
                  shading: { fill: "EAF1F8", type: ShadingType.CLEAR },
                  verticalAlign: VerticalAlign.CENTER,
                  margins: { top: 110, bottom: 110, left: 160, right: 160 },
                  children: [
                    new Paragraph({ children: [new TextRun({ text: k, font: FONT, size: 19, bold: true, color: NAVY })] }),
                  ],
                }),
                new TableCell({
                  width: { size: CONTENT_W - 1900, type: WidthType.DXA },
                  borders: allThin,
                  verticalAlign: VerticalAlign.CENTER,
                  margins: { top: 110, bottom: 110, left: 160, right: 160 },
                  children: [
                    new Paragraph({ children: [new TextRun({ text: v, font: FONT, size: 19, color: "33475B" })] }),
                  ],
                }),
                ,
              ].filter(Boolean),
            })
          ),
        }),
        spacer(280, 0),
        new Paragraph({
          spacing: { before: 120, after: 80 },
          children: [new TextRun({ text: "文档说明", font: FONT, size: 20, bold: true, color: NAVY })],
        }),
        new Paragraph({
          spacing: { after: 120, line: 336, lineRule: LineRuleType.AUTO },
          children: [
            text("本指导书按“拣货—移库—盘点”三条主线梳理横梁货架区手持作业，缺货处理作为拣货的异常分支单列。各章节统一采用“适用场景—操作步骤—现场照片—操作要点”的结构，文中虚线框为照片预留位，可直接替换为现场实拍图，便于培训与现场复核。实际操作界面与按钮名称以手持终端当前版本为准；流程调整时请同步更新本文档及对应照片。"),
          ],
        }),
        spacer(240, 0),
        ...photoSlot("横梁货架区作业现场（照片预留位）", 7600, 3000, "建议使用横向现场照片（16:9）"),
        new Paragraph({ children: [new PageBreak()] }),

        // 1 Overview
        h1("一、作业总览"),
        body("横梁货架区日常作业由拣货、移库、盘点三类任务构成。其中，缺货是拣货过程中的异常分支，异常盘点由缺货单据自动触发。三类作业的触发入口、核心动作与结果输出，见下图与对照表。"),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 100, after: 0, line: 260, lineRule: LineRuleType.AUTO },
          children: [
            new ImageRun({
              type: "png",
              data: fs.readFileSync("/workspace/work/mindmap.png"),
              transformation: { width: 600, height: 364 },
              altText: { title: "横梁货架区作业逻辑导图", description: "拣货、移库、盘点三条作业主线及其子流程", name: "作业逻辑导图" },
            }),
          ],
        }),
        caption("图 1-1　横梁货架区作业逻辑导图"),
        new Paragraph({
          spacing: { before: 60, after: 100 },
          children: [new TextRun({ text: "三类作业对照", font: FONT, size: 20, bold: true, color: NAVY })],
        }),
        new Table({
          width: { size: CONTENT_W, type: WidthType.DXA },
          layout: TableLayoutType.FIXED,
          columnWidths: [1300, 2700, 2800, 2226],
          rows: [
            new TableRow({
              tableHeader: true,
              cantSplit: true,
              children: [
                headCell("作业类型", 1300),
                headCell("操作入口", 2700),
                headCell("核心动作", 2800),
                headCell("结果输出", 2226),
              ],
            }),
            ...overviewRows.map((r, i) =>
              new TableRow({
                cantSplit: true,
                children: r.map((c, j) =>
                  dataCell(
                    c,
                    [1300, 2700, 2800, 2226][j],
                    i % 2 === 0 ? "FFFFFF" : "F4F8FC",
                    j === 0,
                    j === 0
                  )
                ),
              })
            ),
          ],
        }),
        spacer(140, 0),
        body("作业衔接建议：拣货发现库存不足时，先拣完现有数量并提交缺货，由系统生成异常盘点单转入盘点环节；移库完成后必须执行上架确认，确保容器状态闭环。"),

        // 2 Picking
        h1("二、拣货作业"),
        h2("2.1　平置横梁拣货", false),
        body("适用场景：按出库任务在横梁货架区拣取包件；非设备区使用同一拣货功能。"),
        step("pick", "进入功能：", "手持主页面 → 点击【出库】→ 点击【平置横梁拣货】。"),
        step("pick", "选择区域与任务类型：", "区域选择“一层横梁货架区”；任务类型包含 B2B 拣货与加工任务，本次选择【B2B 拣货】。"),
        step("pick", "申领任务：", "在“库位编码”位置扫描库区起始码，连续申领任务，单次最多 15 个；申领完成后点击【执行拣货】。"),
        step("pick", "执行拣货：", "依次扫描拣货库位码、拣货托盘号、拣货书代号；核对书号与库存数量，输入实拣数量，将包件拣至拣货托盘并绑定托盘，完成拣货。"),
        spacer(100, 0),
        photoPair("图 2-1　拣货功能入口界面", "图 2-2　拣货任务申领与执行界面"),
        spacer(160, 0),
        callout("拣货操作要点", [
          "单次申领上限为 15 个，避免一次申领过多造成任务积压。",
          "拣货前核对书号与库存数量，数量以实物清点结果为准。",
          "包件必须与拣货托盘绑定，未绑定的包件不得进入后续环节。",
        ]),

        h2("2.2　出库绑箱"),
        body("适用场景：拣货完成后，将拣货托盘内的包件绑定至周转箱并投线。"),
        step("box", "扫描拣货托盘号：", "在出库绑箱页面扫描待绑定的拣货托盘号。"),
        step("box", "扫描管理书代号：", "扫描托盘内包件的管理书代号，确认书号与托盘对应。"),
        step("box", "核对数量：", "核对任务数与实拣数量，确认二者一致。"),
        step("box", "绑定周转箱：", "扫描目标容器号，完成周转箱绑定。"),
        step("box", "完成投线：", "绑定完成后即可投线。"),
        spacer(100, 0),
        ...photoSlot("图 2-3　出库绑箱界面"),
        spacer(160, 0),
        callout("出库绑箱要点", [
          "任务数与实拣数量不一致时不得绑箱，须先查清差异。",
          "托盘号与周转箱号须逐一扫码绑定，避免混箱。",
        ]),

        h2("2.3　缺货处理（拣货异常分支）"),
        body("适用场景：拣货时库位库存不足，无法完成全部任务数量。"),
        step("short", "先拣现有数量：", "库存不足时，先拣完库位现有数量。"),
        step("short", "提交缺货：", "剩余数量在拣货页面点击【缺货】。"),
        step("short", "确认结果：", "系统提示“缺货成功”。"),
        step("short", "库位管控：", "该库位自动进入管控状态。"),
        step("short", "生成异常盘点单：", "系统自动生成此库位的异常盘点单，后续按盘点单处理（见 4.2）。"),
        spacer(100, 0),
        photoPair("图 2-4　缺货提交与提示界面", "图 2-5　系统生成的异常盘点单"),
        spacer(160, 0),
        callout("缺货处理要点", [
          "缺货提交后库位即被管控，不可继续按原任务拣货。",
          "异常盘点单由系统自动生成，无需线下另行开单。",
        ]),

        // 3 Move
        h1("三、移库作业"),
        h2("3.1　分段移库（下架 → 上架确认）", false),
        body("适用场景：横梁货架区库位之间需分段转移，或将容器移至目标区域后重新上架。"),
        new Paragraph({
          spacing: { before: 80, after: 80 },
          children: [new TextRun({ text: "第一步　下架", font: FONT, size: 20, bold: true, color: "24425F" })],
        }),
        step("move", "进入功能：", "手持主页面 → 点击【库存】→ 找到【分段移库下架】。"),
        step("move", "选择区域：", "下架区域选择“横梁货架区”。"),
        step("move", "查询待移库位：", "点击“容器号”旁的放大镜，查看需要移库的库位书代号及目标区域。"),
        step("move", "确认下架：", "核对库位容器号并输入书代号，确认下架。"),
        new Paragraph({
          spacing: { before: 160, after: 80 },
          children: [new TextRun({ text: "第二步　上架确认", font: FONT, size: 20, bold: true, color: "24425F" })],
        }),
        step("putaway", "到达目标区域：", "手持主页面 → 点击【入库】→ 点击【上架确认】。"),
        step("putaway", "录入容器信息：", "输入容器号、管理书代号，核对库存数量。"),
        step("putaway", "确认上架：", "扫描目标库位，确认上架成功。"),
        spacer(100, 0),
        photoPair("图 3-1　分段移库下架界面", "图 3-2　上架确认界面"),
        spacer(160, 0),
        callout("分段移库操作要点", [
          "下架后须在目标区域完成上架确认，避免形成“已下架、未上架”的在途容器。",
          "容器号与书代号须与下架记录一致，数量核对后再确认。",
        ]),

        h2("3.2　标准移库"),
        body("适用场景：同一区域内库位到库位的直接移库，无需分段下架、上架。"),
        step("std", "进入功能：", "手持主页面 → 点击【库存】→ 选择【标准移库】。"),
        step("std", "核对原始信息：", "扫描原始库位、原始容器号，核对信息后确认。"),
        step("std", "输入移库数量：", "按实际转移数量录入。"),
        step("std", "扫描目标信息：", "扫描目标库位、目标容器号。"),
        step("std", "完成移库：", "点击确认，完成标准移库。"),
        spacer(100, 0),
        ...photoSlot("图 3-3　标准移库界面"),
        spacer(160, 0),
        callout("标准移库操作要点", [
          "移库数量不得超过原库位可用库存。",
          "目标库位与目标容器号须成对扫描，确认对应关系无误。",
        ]),

        // 4 Count
        h1("四、盘点作业"),
        h2("4.1　非设备区盘点", false),
        body("适用场景：横梁货架区（非设备区）日常盘点。"),
        step("count", "进入功能：", "手持主页面 → 点击【库存】→ 点击【非设备区盘点】。"),
        step("count", "选择区域：", "区域选择“一层横梁货架区”，扫描起始码。"),
        step("count", "接收推荐库位：", "系统自动推荐盘点库位，点击确认。"),
        step("count", "进入盘点页面：", "扫描盘点库位码，输入托盘号、管理书代号。"),
        step("count", "清点实物：", "清点实物数量，输入准确数量。"),
        step("count", "提交结果：", "点击确认，提示“盘点成功”。"),
        spacer(100, 0),
        photoPair("图 4-1　盘点入口与起始码扫描", "图 4-2　盘点数量录入界面"),
        spacer(160, 0),
        callout("盘点操作要点", [
          "盘点数量以实物清点为准，扫码与录入信息须与库位实际一致。",
          "盘点确认后结果直接生效，提交前应再次核对数量。",
        ]),

        h2("4.2　异常盘点（缺货触发）"),
        body("适用场景：拣货提交缺货后，系统自动生成该库位的异常盘点单。"),
        step("abn", "查看单据：", "缺货成功后，系统自动生成该库位的异常盘点单。"),
        step("abn", "执行盘点：", "按盘点单逐项核对库位实物与账面数量。"),
        step("abn", "处理差异：", "盘点结果以盘点单为准，确认后按盘点结果更新库存。"),
        spacer(100, 0),
        ...photoSlot("图 4-3　异常盘点单"),
        spacer(160, 0),
        callout("异常盘点要点", [
          "异常盘点单由系统自动生成，无需人工新建。",
          "缺货与异常盘点在同一库位形成闭环，处理完成后方可恢复正常作业。",
        ]),

        // 5 Rules
        h1("五、通用操作要点"),
        new Table({
          width: { size: CONTENT_W, type: WidthType.DXA },
          layout: TableLayoutType.FIXED,
          columnWidths: [700, 1600, 6726],
          rows: [
            new TableRow({
              tableHeader: true,
              cantSplit: true,
              children: [
                headCell("序号", 700),
                headCell("要点", 1600),
                headCell("具体要求", 6726),
              ],
            }),
            ...rulesRows.map((r, i) =>
              new TableRow({
                cantSplit: true,
                children: r.map((c, j) =>
                  dataCell(c, [700, 1600, 6726][j], i % 2 === 0 ? "FFFFFF" : "F4F8FC", j < 2, j === 1)
                )
              })
            ),
          ],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("/workspace/横梁货架区手持作业指导书.docx", buf);
  console.log("done");
});
