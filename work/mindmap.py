import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from matplotlib.path import Path
import matplotlib.patches as mpatches

plt.rcParams["font.family"] = ["Noto Sans CJK SC"]

W, H = 18.0, 10.0
fig = plt.figure(figsize=(10, 10 * H / W), dpi=180)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, W)
ax.set_ylim(0, H)
ax.axis("off")
fig.patch.set_facecolor("white")

C_TEXT = "#2B2B2B"
C_GRAY = "#6B6B6B"

def box(x0, y0, w, h, fill, edge, lw=1.6, r=0.12, z=3):
    p = FancyBboxPatch((x0, y0), w, h,
                       boxstyle=f"round,pad=0.02,rounding_size={r}",
                       linewidth=lw, edgecolor=edge, facecolor=fill, zorder=z)
    ax.add_patch(p)
    return p

def label(cx, cy, text, size, color=C_TEXT, weight="normal", z=4):
    ax.text(cx, cy, text, ha="center", va="center", fontsize=size, color=color,
            fontweight=weight, zorder=z, linespacing=1.35)

def curve(x1, y1, x2, y2, color, lw=1.8, z=2):
    dx = (x2 - x1) * 0.5
    verts = [(x1, y1), (x1 + dx * 0.55, y1), (x1 + dx * 1.45, y2), (x2, y2)]
    codes = [Path.MOVETO, Path.CURVE4, Path.CURVE4, Path.CURVE4]
    ax.add_patch(mpatches.PathPatch(Path(verts, codes), edgecolor=color,
                                    facecolor="none", lw=lw, zorder=z, capstyle="round"))

# root
RX, RY, RW, RH = 0.30, 4.05, 2.40, 1.35
box(RX, RY, RW, RH, "#F4B183", "#C55A11", 2.0, 0.16)
label(RX + RW / 2, RY + RH / 2, "横梁货架区\n手持作业", 15, "#7B3F00", "bold")

branches = [
    dict(name="拣货作业", y=7.75, fill="#B7E1E4", edge="#3F939C", text="#1F5E64",
         kids=[("平置横梁拣货", "出库任务拣货，非设备区共用", 8.55),
               ("缺货处理（异常）", "库存不足时触发", 7.75),
               ("出库绑箱", "绑箱完成后投线", 6.95)]),
    dict(name="移库作业", y=4.65, fill="#CFC9F0", edge="#6055A5", text="#3B3273",
         kids=[("分段移库下架", "查询待移库位并下架", 5.45),
               ("上架确认", "目标区域入库上架", 4.65),
               ("标准移库", "库位到库位直接移库", 3.85)]),
    dict(name="盘点作业", y=1.9, fill="#F7E6A6", edge="#B8912A", text="#6B5310",
         kids=[("非设备区盘点", "系统推荐盘点库位", 2.4),
               ("异常盘点（缺货触发）", "缺货自动生成盘点单", 1.4)]),
]

L1X, L1W, L1H = 3.05, 2.35, 0.95
L2X, L2W, L2H = 6.25, 3.75, 0.66
AX = 10.2

for b in branches:
    cy = b["y"]
    curve(RX + RW, RY + RH / 2, L1X, cy, b["edge"], 2.2)
    box(L1X, cy - L1H / 2, L1W, L1H, b["fill"], b["edge"], 1.8, 0.12)
    label(L1X + L1W / 2, cy, b["name"], 16, b["text"], "bold")

    for txt, note, ky in b["kids"]:
        curve(L1X + L1W, cy, L2X, ky, b["edge"], 1.7)
        box(L2X, ky - L2H / 2, L2W, L2H, "#FFFFFF", b["edge"], 1.2, 0.09)
        # colored left accent inside child
        box(L2X + 0.02, ky - L2H / 2 + 0.06, 0.09, L2H - 0.12, b["fill"], b["fill"], 0, 0.03)
        label(L2X + L2W / 2 + 0.05, ky, txt, 13, C_TEXT, "bold")
        ax.text(AX, ky, note, ha="left", va="center", fontsize=11,
                color=C_GRAY, zorder=4)

# bottom common-rules band
BX, BY, BW, BH = 0.35, 0.18, 14.55, 0.72
box(BX, BY, BW, BH, "#F4F6F8", "#C8D0D6", 1.2, 0.10)
ax.text(BX + 0.3, BY + BH / 2, "通用要点", ha="left", va="center",
        fontsize=12, color="#1F4E79", fontweight="bold", zorder=5)
ax.text(BX + 1.62, BY + BH / 2,
        "单次申领 ≤ 15 单   ·   书号 / 数量 / 容器号 三重核对   ·   扫码核对后再确认   ·   异常全部系统留痕",
        ha="left", va="center", fontsize=10.5, color=C_GRAY, zorder=5)

fig.savefig("/workspace/work/mindmap.png", dpi=180, facecolor="white", bbox_inches="tight", pad_inches=0.18)

from PIL import Image, ImageChops
im = Image.open("/workspace/work/mindmap.png").convert("RGB")
bg = Image.new("RGB", im.size, "white")
bb = ImageChops.difference(im, bg).getbbox()
pad = 18
box = (max(0, bb[0]-pad), max(0, bb[1]-pad), min(im.size[0], bb[2]+pad), min(im.size[1], bb[3]+pad))
im.crop(box).save("/workspace/work/mindmap.png")
print("cropped", box, im.size)
print("saved")
