# -*- coding: utf-8 -*-
"""生成《多平台商品上架字段总表_马洁填写版》"""
try:
    import openpyxl
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.page import PageMargins

OUT = "/Users/lhwylp/Desktop/ai-work/pri-ws/majie/repo/majie/多平台商品上架字段总表_马洁填写版.xlsx"
TITLE = "多平台商品上架字段总表（马洁填写版）"


def xl_color(css_hex: str) -> str:
    v = css_hex.removeprefix("#").upper()
    if len(v) != 6:
        raise ValueError(f"Expected #RRGGBB, got: {css_hex}")
    return "FF" + v


XL_PRIMARY = xl_color("#4472C4")
XL_LIGHT = xl_color("#D9E2F3")
XL_INPUT = xl_color("#FFF2CC")
XL_CALC = xl_color("#F2F2F2")
XL_BORDER = xl_color("#BFBFBF")
XL_WARN = xl_color("#FCE4D6")
XL_GREEN_F = xl_color("#C6EFCE")
XL_GREEN_T = xl_color("#006100")
XL_RED_F = xl_color("#FFC7CE")
XL_RED_T = xl_color("#9C0006")
XL_YEL_F = xl_color("#FFEB9C")
XL_YEL_T = xl_color("#9C6500")
XL_GRAY_T = xl_color("#808080")
XL_WHITE = xl_color("#FFFFFF")

thin = Side(style="thin", color=XL_BORDER)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
FONT = "微软雅黑"

wb = Workbook()
wb.properties.title = TITLE


def style_range(ws, cell_range, font=None, fill=None, align=None, border=True, fmt=None):
    for row in ws[cell_range]:
        for c in row:
            if font:
                c.font = font
            if fill:
                c.fill = fill
            if align:
                c.alignment = align
            if border:
                c.border = BORDER
            if fmt:
                c.number_format = fmt


def title_row(ws, text, ncols, row=1, sub=None):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=15, bold=True, color=XL_WHITE)
    c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 34
    if sub:
        ws.merge_cells(start_row=row + 1, start_column=1, end_row=row + 1, end_column=ncols)
        c2 = ws.cell(row=row + 1, column=1, value=sub)
        c2.font = Font(name=FONT, size=10, color="FF595959")
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
        ws.row_dimensions[row + 1].height = 26


def header_row(ws, row, headers, widths=None, height=32):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name=FONT, size=10, bold=True, color=XL_WHITE)
        c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = height
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


# ============================================================
# 01_填写说明
# ============================================================
ws = wb.active
ws.title = "01_填写说明"
ws.sheet_properties.tabColor = XL_PRIMARY
headers = ["工作表", "谁填", "要做什么", "说明"]
widths = [22, 14, 46, 62]
title_row(ws, "多平台商品上架字段总表 · 填写说明", 4,
          sub="适用：汽车贴膜店（玻璃膜 / 隐形车衣 / 改色膜 / 内饰改装 / 音响隔音）。覆盖平台：美团·大众点评、抖音来客、快手本地生活、小红书、闲鱼、淘宝。")
header_row(ws, 3, headers, widths)

rows = [
    ("02_门店通用信息", "马洁", "填一次，全平台复用", "门店主体、地址、电话、资质、结算账户、履约口径。这些在所有平台入驻时都会再问一遍，统一口径才不会反复被驳回。"),
    ("03_套餐填写表", "马洁", "★核心：一个套餐一列，从上往下填", "80 个字段按 7 大类排好，E~J 六列对应六个套餐。黄色格子就是你要填的。"),
    ("04_平台必填对照", "不用填", "查：这个字段哪个平台必须填", "字段 × 平台矩阵。绿色底=必填，白底=选填，“— ”=该平台没有这个字段。"),
    ("05_平台字段明细", "不用填", "查：某个平台后台长什么样", "按平台分块，写清字段类型、可选值、字数/图片限制、易驳回点。上架时对着这一块填。"),
    ("06_素材清单", "马洁", "勾：素材有没有、在哪", "拍什么、要几张、什么比例，一次拍全，后面六个平台都能用。"),
    ("07_合规与话术", "不用填", "查：什么不能写、标题怎么写", "禁词、必须写清的规则、标题公式、私信首响话术。上架前过一遍。"),
    ("08_发布进度", "运营", "跟：每个平台每个套餐走到哪了", "含自动统计。被驳回的把原因填进去，改完再提交。"),
]
r = 4
for row in rows:
    fill = PatternFill("solid", fgColor=XL_LIGHT) if str(row[1]) == "马洁" else None
    for i, v in enumerate(row, start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.font = Font(name=FONT, size=10, bold=(i == 1))
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True,
                                indent=0)
        c.border = BORDER
        if fill:
            c.fill = fill
    ws.row_dimensions[r].height = 34
    r += 1

# 颜色约定
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
c = ws.cell(row=r, column=1, value="颜色约定")
c.font = Font(name=FONT, size=11, bold=True, color=XL_WHITE)
c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[r].height = 24
r += 1
legend = [
    ("浅黄色格", XL_INPUT, "你要填的"),
    ("浅灰色格", XL_CALC, "自动算的，别动"),
    ("浅蓝色行", XL_LIGHT, "分组标题/给马洁的表"),
]
for name, color, desc in legend:
    ws.cell(row=r, column=1, value=name).fill = PatternFill("solid", fgColor=color)
    ws.cell(row=r, column=1).font = Font(name=FONT, size=10)
    ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=2, value=desc).font = Font(name=FONT, size=10)
    ws.cell(row=r, column=2).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    r += 1

# 三条原则
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
c = ws.cell(row=r, column=1, value="开工前先定这 3 条，能省掉一半返工")
c.font = Font(name=FONT, size=11, bold=True, color=XL_WHITE)
c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[r].height = 24
r += 1
principles = [
    ("1. 价格只有一套口径",
     "门市价必须真实可查（别为了显得便宜写 9980 现价 1999），团购价在各平台保持一致。平台之间差几块钱没关系，差几百会被用户截图对比。"),
    ("2. 规则要写在用户看得见的地方",
     "有效期、能不能用节假日、要不要预约、大车加不加价、退款怎么退 —— 这五条不写清，退款和差评基本都从这里来。写清楚了，抖音/快手审核也过得快。"),
    ("3. 图片一次拍全，六平台复用",
     "每个平台都要的：门头、环境、施工过程、前后对比。只有小红书要 3:4 竖图、抖音/快手要 9:16 视频、美团要 16:9 头图。按最严的规格拍，往下裁就行。"),
]
for name, desc in principles:
    ws.cell(row=r, column=1, value=name).font = Font(name=FONT, size=10, bold=True)
    ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    c = ws.cell(row=r, column=2, value=desc)
    c.font = Font(name=FONT, size=10)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    for col in range(1, 5):
        ws.cell(row=r, column=col).border = BORDER
    ws.row_dimensions[r].height = 42
    r += 1

# 示例
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
c = ws.cell(row=r, column=1, value="03 表长这样（示例：只填了一个套餐）")
c.font = Font(name=FONT, size=11, bold=True, color=XL_WHITE)
c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[r].height = 24
r += 1
ex_head = ["字段", "填写要求", "套餐2 隐形车衣·TPU主推"]
for i, h in enumerate(ex_head, start=1):
    c = ws.cell(row=r, column=i, value=h)
    c.font = Font(name=FONT, size=10, bold=True, color=XL_WHITE)
    c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
r += 1
examples = [
    ("套餐名称", "品牌+项目+规格+质保，30字内", "卡皮诺 隐形车衣 全车 5年质保"),
    ("团购价（元）", "实际成交价", 2680),
    ("门市价（元）", "原价，必须真实", 6680),
    ("质保年限", "3年/5年/10年", "5年"),
    ("施工时长（小时）", "含交车等待", 24),
]
for name, req, val in examples:
    ws.cell(row=r, column=1, value=name).font = Font(name=FONT, size=10, bold=True)
    ws.cell(row=r, column=2, value=req).font = Font(name=FONT, size=10, color="FF595959")
    c = ws.cell(row=r, column=3, value=val)
    c.fill = PatternFill("solid", fgColor=XL_INPUT)
    c.font = Font(name=FONT, size=10)
    for col in range(1, 4):
        ws.cell(row=r, column=col).border = BORDER
        ws.cell(row=r, column=col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    if isinstance(val, int):
        ws.cell(row=r, column=3).number_format = "¥#,##0"
    r += 1

# 置信度
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
c = ws.cell(row=r, column=1, value="资料来源与可信度（重要）")
c.font = Font(name=FONT, size=11, bold=True, color=XL_WHITE)
c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[r].height = 24
r += 1
src = [
    ("美团 / 大众点评", "高", "来自经营宝后台逐字段记录（车衣、改色膜两份完整清单已在仓库），可直接照着填。大众点评与美团共用同一套经营宝后台，字段一致。"),
    ("抖音来客", "较高", "来自后台创建商品的实际流程 + 官方商品发布规则。个别字段名各版本略有差异（如“建议用餐人数”对贴膜店填车辆数）。"),
    ("快手本地生活", "较高", "来自快手学苑《商品发布规则与常见问题》+ 商家后台流程。快手对图片要求最严（禁黑边、禁截图、禁第三方水印）。"),
    ("小红书", "中", "企业专业号 + POI 认领 + 本地生活团购的路径是确定的；“汽车服务/洗车养护”类目能否自助开通，不同资料说法不一，需在后台或服务商处确认（见 05 表标注）。"),
    ("闲鱼", "中", "走“发服务”入口，字段以公开流程整理；具体类目树随版本变动，发布时以实际下拉为准。"),
    ("淘宝", "中", "走“服务 > 汽车服务”类目需资质（执照经营范围含汽车美容）；类目属性随类目不同，上架时以系统必填项为准。"),
]
for name, conf, desc in src:
    ws.cell(row=r, column=1, value=name).font = Font(name=FONT, size=10, bold=True)
    ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
    cc = ws.cell(row=r, column=2, value=conf)
    cc.alignment = Alignment(horizontal="center", vertical="top")
    cc.font = Font(name=FONT, size=10, bold=True,
                   color=XL_GREEN_T if conf == "高" else (XL_YEL_T if conf == "较高" else XL_RED_T))
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    c = ws.cell(row=r, column=3, value=desc)
    c.font = Font(name=FONT, size=10)
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
    for col in range(1, 5):
        ws.cell(row=r, column=col).border = BORDER
    ws.row_dimensions[r].height = 46
    r += 1

ws.freeze_panes = "A4"
ws.sheet_view.showGridLines = False


# ============================================================
# 02_门店通用信息
# ============================================================
ws = wb.create_sheet("02_门店通用信息")
ws.sheet_properties.tabColor = XL_INPUT
headers = ["分类", "字段", "必填", "填写要求 / 示例", "填写内容", "哪个平台要", "备注"]
widths = [14, 22, 8, 48, 30, 22, 26]
title_row(ws, "门店通用信息（填一次，六个平台复用）", 7,
          sub="这一页的信息在每家平台入驻/门店认领时都会被问到。先在这里定好口径，后面照抄，避免“地址不一致被驳回”这类问题。")
header_row(ws, 3, headers, widths)

store_rows = [
    ("门店主体", "门店对外展示名", "必填", "与门头照一致，12 字以内更稳", "津豪汽车贴膜改装", "全部", ""),
    ("门店主体", "营业执照名称", "必填", "与执照完全一致，一个字都不能差", "", "全部", ""),
    ("门店主体", "统一社会信用代码", "必填", "18 位", "", "全部", ""),
    ("门店主体", "主体类型", "必填", "个体工商户 / 有限公司", "", "全部", "个体工商户在部分高风险类目会受限"),
    ("门店主体", "法人姓名", "必填", "与执照一致", "", "全部", ""),
    ("门店主体", "法人手机号", "必填", "能接收验证码、能接审核电话", "", "全部", ""),
    ("门店主体", "执照注册地址", "必填", "与执照完全一致", "", "全部", ""),
    ("门店主体", "实际经营地址", "必填", "写到门牌号/楼层，便于审核比对", "", "全部", "与执照不一致时，备好一句解释口径"),
    ("门店主体", "地图定位名称（POI）", "必填", "地图上能搜到的名称；搜不到就先新建", "", "抖音 快手 小红书 美团 点评", "POI 名称与门头照不一致是高频驳回原因"),
    ("门店主体", "经纬度", "选填", "地图拾取后填，便于新建 POI", "", "全部", ""),
    ("联系方式", "门店电话（主要）", "必填", "营业时间必须有人接", "13821425827", "全部", ""),
    ("联系方式", "门店电话（备用）", "选填", "老板手机", "13820908557", "全部", ""),
    ("联系方式", "微信号/私信口径", "选填", "平台内优先，不主动引导站外", "", "全部", "图片/文案里出现微信号会被限流"),
    ("经营信息", "营业时间", "必填", "含午休，例：09:00-18:00（午休 12:00-13:00）", "9:00-18:00", "全部", ""),
    ("经营信息", "每周店休", "必填", "无休 / 周日休", "", "全部", "影响“可用时段”怎么填"),
    ("经营信息", "门店类目（所在行业）", "必填", "汽车服务 > 汽车美容/贴膜；美团为 爱车 > 美容", "", "全部", "类目填错，后面流量全偏"),
    ("经营信息", "门店简介（短·50字）", "必填", "做什么 + 做了多久 + 一句信任点", "", "全部", ""),
    ("经营信息", "门店简介（长·200字）", "选填", "适合美团/大众点评商户介绍", "", "美团 点评 小红书", ""),
    ("经营信息", "工位/施工车间数量", "必填", "数字，决定日产能", "", "内部", ""),
    ("经营信息", "日均最大接单量", "必填", "台/天，决定库存怎么设", "", "内部", ""),
    ("经营信息", "可服务车型范围", "必填", "轿车/SUV/MPV/新能源；是否接超跑、改装车", "", "全部", ""),
    ("经营信息", "交通与停车指引", "选填", "地铁/公交 + 停车是否免费", "", "美团 点评 小红书", "写清楚能减少“找不到店”的退款"),
    ("资质证照", "营业执照照片", "必填", "清晰、四角完整、不反光", "", "全部", ""),
    ("资质证照", "法人身份证正反面", "必填", "清晰、不遮挡", "", "全部", ""),
    ("资质证照", "法人手持身份证照", "选填", "部分平台人脸核验时代替", "", "部分平台", ""),
    ("资质证照", "品牌授权书", "选填", "雷朋/3M/科为等，有就传，能提审核通过率", "", "美团 抖音 淘宝", ""),
    ("资质证照", "行业许可证", "选填", "贴膜一般不需要；改内饰/改装按当地要求", "", "按类目", ""),
    ("资质证照", "门头照", "必填", "横版、清晰、招牌完整、无遮挡", "", "全部", "用网图会被直接驳回"),
    ("资质证照", "店内环境照", "必填", "工位、车间、休息区各 1-2 张", "", "全部", ""),
    ("结算", "结算账户类型", "必填", "对公 / 对私（个体户常见）", "", "全部", ""),
    ("结算", "开户行 + 卡号", "必填", "与主体一致", "", "全部", "主体不一致无法结算"),
    ("结算", "发票能力", "选填", "能否开票、票种", "", "淘宝 抖音", ""),
    ("履约与售后", "预约方式", "必填", "电话预约 / 平台预约 / 私信预约", "", "全部", ""),
    ("履约与售后", "施工时长口径", "必填", "玻璃膜 2-4h；车衣 1-2 天；改色 2-3 天；内饰 1-3 天", "", "全部", "影响“是否需要预约”和库存"),
    ("履约与售后", "能否留车过夜", "必填", "能 / 不能；能否代步车", "", "全部", "车衣改色必答，否则差评多"),
    ("履约与售后", "改期/取消规则", "必填", "例：提前 24h 免费改期，最多 2 次", "", "全部", ""),
    ("履约与售后", "质保范围", "必填", "例：起泡、脱胶、开裂、明显褪色；不含人为剐蹭", "", "全部", "这是用户最看重的差异点"),
    ("履约与售后", "质保凭证形式", "必填", "电子质保卡 / 纸质卡 + 编号", "", "全部", ""),
    ("履约与售后", "售后响应时效", "必填", "例：24h 内响应，7 天内免费返工", "", "全部", ""),
    ("履约与售后", "售后需客户提供的证据", "必填", "施工前车况照片、质保卡编号", "", "内部", ""),
    ("线上资产", "抖音号（企业号）", "选填", "已认证蓝V填这里", "", "抖音", ""),
    ("线上资产", "小红书专业号", "选填", "已认证填这里", "", "小红书", ""),
    ("线上资产", "美团/点评门店 ID", "选填", "认领后后台可见", "", "美团 点评", ""),
    ("线上资产", "闲鱼账号", "选填", "建议专号专用", "", "闲鱼", ""),
    ("线上资产", "淘宝店铺名", "选填", "已开店填这里", "", "淘宝", ""),
]

r = 4
prev_cat = None
for row in store_rows:
    cat = row[0]
    if cat != prev_cat:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        c = ws.cell(row=r, column=1, value=cat)
        c.font = Font(name=FONT, size=10, bold=True, color="FF1F3864")
        c.fill = PatternFill("solid", fgColor=XL_LIGHT)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in range(1, 8):
            ws.cell(row=r, column=col).border = BORDER
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=XL_LIGHT)
        ws.row_dimensions[r].height = 22
        r += 1
        prev_cat = cat
    for i, v in enumerate(row, start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.border = BORDER
        c.font = Font(name=FONT, size=10, bold=(i == 2))
        if i == 5:
            c.fill = PatternFill("solid", fgColor=XL_INPUT)
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        elif i in (3, 6):
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        else:
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    ws.row_dimensions[r].height = 26
    r += 1

ws.freeze_panes = "C4"
ws.auto_filter.ref = f"A3:G{r-1}"


# ============================================================
# 03_套餐填写表（核心）
# ============================================================
ws = wb.create_sheet("03_套餐填写表")
ws.sheet_properties.tabColor = xl_color("#C00000")
PKGS = ["套餐1\n玻璃隔热膜", "套餐2\n隐形车衣·TPU主推", "套餐3\n隐形车衣·PVC引流",
        "套餐4\n车身改色膜", "套餐5\n内饰改装·座椅包覆", "套餐6\n音响·隔音升级"]
headers = ["分类", "字段", "填写要求 / 示例", "哪个平台要"] + PKGS + ["备注"]
widths = [13, 21, 44, 20] + [19] * 6 + [22]
title_row(ws, "套餐填写表 · 一个套餐一列，从上往下填", 11,
          sub="黄色格＝你填；灰色格＝自动算。填不出来的先写“待定”，别空着——空着表就白做了。同类字段六个套餐都要填，不必填满六列，做几个套餐填几列。")
header_row(ws, 3, headers, widths, height=40)

# (分类, 字段, 填写要求, 覆盖平台, kind, fmt, key)
F = []
add = lambda *a: F.append(a)

# A 基础信息
add("A 基础信息", "套餐编号（内部）", "自定义短码，如 JHM-01 / JHM-02", "内部", "in", None, "code")
add("A 基础信息", "套餐名称", "品牌+项目+规格+质保，30 字内。例：卡皮诺 隐形车衣 全车 5年质保", "全部", "in", None, "name")
add("A 基础信息", "商品标题·抖音", "【品牌】+项目+规格+车型，30 字内。例：【卡皮诺】隐形车衣全车 TPU 5年质保", "抖音", "in", None, "dy_title")
add("A 基础信息", "商品标题·快手", "门店品牌简称+套餐名+规格；多店通用要注明。例：津豪贴膜 隐形车衣全车 5年质保", "快手", "in", None, "ks_title")
add("A 基础信息", "商品标题·闲鱼", "前 15 字放核心词。例：天津汽车贴膜 隐形车衣全车 TPU 5年质保", "闲鱼", "in", None, "xy_title")
add("A 基础信息", "商品标题·淘宝", "项目词+品牌+规格+质保，30 字内堆关键词，别堆重复词", "淘宝", "in", None, "tb_title")
add("A 基础信息", "商品类型", "团购套餐 / 代金券 / 次卡（贴膜一般选团购套餐）", "美团 抖音 快手 小红书", "in", None, "ptype")
add("A 基础信息", "平台类目路径", "每个平台一条，路径见 05 表。例：抖音 汽车服务>汽车美容>贴膜", "全部", "in", None, "cat")
add("A 基础信息", "一句话卖点", "≤15 字，能当封面大字。例：防爆隔热 · 5 年质保", "全部", "in", None, "hook")
add("A 基础信息", "商品短描述", "50 字内，列表页/卡片用：包含什么 + 谁适合 + 一句保障", "全部", "in", None, "short")
add("A 基础信息", "图文详情文案", "≤1000 字。四段：包含什么 / 怎么施工 / 质保到什么程度 / 注意事项", "美团 抖音 淘宝 小红书", "in", None, "detail")
add("A 基础信息", "特色标签", "≤6 字，如：防刮提亮 / 隔热降温 / 不留胶", "美团 抖音", "in", None, "tag")
add("A 基础信息", "适用车型", "全车型 / 五座轿车 / SUV·MPV；有加价的写清哪些车型加多少", "全部", "in", None, "car")
add("A 基础信息", "适用位置", "全车 / 前挡 / 侧后挡 / 单座椅 / 局部（如大灯、天窗）", "全部", "in", None, "pos")
add("A 基础信息", "份数 / 次数", "1 份；按次卖的写清几次（如内饰清洁 2 次）", "全部", "in", None, "qty")

# B 属性
add("B 商品属性", "膜品牌", "雷朋 / 3M / 科为 / 卡皮诺 等，写全称", "全部", "in", None, "brand")
add("B 商品属性", "型号 / 系列", "具体型号，如 雷朋 LX-70、卡皮诺 CP-5", "全部", "in", None, "model")
add("B 商品属性", "材质", "TPU / PVC / 陶瓷 / 磁控金属 / 纳尔；玻璃膜写清工艺", "美团 抖音 淘宝", "in", None, "material")
add("B 商品属性", "净膜厚度（mil）", "数字，如 7.5。车衣类必填，薄厚是用户比价重点", "美团 抖音", "in", None, "thick")
add("B 商品属性", "修复类型", "自修复 / 热修复 / 无（车衣专属字段）", "美团", "in", None, "heal")
add("B 商品属性", "表面质感", "光面 / 雾面 / 哑光 / 电镀 / 碳纤纹", "美团", "in", None, "surface")
add("B 商品属性", "耐变黄时间", "5 年 / 10 年（车衣）", "美团", "in", None, "yellow")
add("B 商品属性", "生产工艺", "压延 / 铸造（改色膜）", "美团", "in", None, "craft")
add("B 商品属性", "颜色 / 图案", "改色膜填颜色名；每个颜色配一张色卡图，最多 5 组", "美团 淘宝", "in", None, "color")
add("B 商品属性", "可选颜色数", "数字，如 12 色可选", "改色", "in", None, "colorn")
add("B 商品属性", "红外阻隔率（%）", "数字，如 95。玻璃膜核心参数", "玻璃膜", "in", "0", "ir")
add("B 商品属性", "紫外线阻隔率（%）", "数字，如 99", "玻璃膜", "in", "0", "uv")
add("B 商品属性", "可见光透光率 VLT（%）", "数字，如 70。前挡按法规需 ≥70%，别写错", "玻璃膜", "in", "0", "vlt")
add("B 商品属性", "质保年限", "3 年 / 5 年 / 10 年", "全部", "in", None, "warranty")
add("B 商品属性", "质保范围", "逐条写：起泡 / 脱胶 / 开裂 / 明显褪色；不含人为剐蹭、事故损伤", "全部", "in", None, "wscope")
add("B 商品属性", "生产厂商 / 产地", "文本，如 卡皮诺（国产）/ 雷朋（日本）", "美团 淘宝", "in", None, "maker")
add("B 商品属性", "施工流程", "清洁→验车记录→专车裁膜→施工→质检→交车；多选写入", "美团 抖音 快手", "in", None, "flow")
add("B 商品属性", "增值服务", "赠车衣险 / 免费修补 N 次 / 送洗车 / 上门取送车 / 旧膜免费撕除", "全部", "in", None, "extra")
add("B 商品属性", "建议车辆数", "抖音“建议人数”一栏对贴膜店填 1 车；多车套餐才填 2 车", "抖音", "in", "0", "num")

# C 价格
add("C 价格与库存", "门市价（元）", "原价，必须真实可查。别虚高再打折", "全部", "in", "¥#,##0.00", "list_price")
add("C 价格与库存", "团购价 / 售价（元）", "实际成交价，各平台保持一致（允许几元差）", "全部", "in", "¥#,##0.00", "price")
add("C 价格与库存", "材料成本（元）", "内部核算用，不对外；膜料+辅料", "内部", "in", "¥#,##0.00", "cost")
add("C 价格与库存", "预估结算价（元）", "自动：团购价 × 94%（美团基础服务费 6%）", "计算", "calc", "¥#,##0.00", "settle")
add("C 价格与库存", "预估毛利（元）", "自动：结算价 − 材料成本（未含工时/房租）", "计算", "calc", "¥#,##0.00", "gross")
add("C 价格与库存", "毛利率", "自动：毛利 ÷ 结算价", "计算", "calc", "0.0%", "margin")
add("C 价格与库存", "是否含施工费", "含 / 不含（不含要写清另收多少）", "全部", "in", None, "inlabor")
add("C 价格与库存", "库存 / 每日可接单", "数字或不限；建议设为“日均产能 × 7”，别设 2000000", "美团 抖音 快手 淘宝", "in", "0", "stock")
add("C 价格与库存", "每人限购", "不限制 / N 份", "全部", "in", None, "limit")
add("C 价格与库存", "大车 / SUV 加价", "例：SUV +200，MPV +300，超跑另议；不加价就写“不加价”", "全部", "in", None, "surge")
add("C 价格与库存", "促销 / 立减", "例：新客立减 50；没有就写“无”", "美团 抖音", "in", None, "promo")

# D 规则
add("D 使用规则", "有效期", "购买后 N 天内有效（建议 90 天）或固定起止日", "全部", "in", None, "valid")
add("D 使用规则", "可用时段", "营业时间内，例：09:00-18:00；车衣需留车则注明“需预约整天空档”", "全部", "in", None, "hours")
add("D 使用规则", "不可用日期", "法定节假日是否可用；春节前后是否停接", "全部", "in", None, "blackout")
add("D 使用规则", "是否需预约", "需提前 1 天 / 需提前 3 天（车衣）/ 无需预约", "全部", "in", None, "book")
add("D 使用规则", "预约方式", "电话 / 平台内私信 / 平台预约组件", "全部", "in", None, "bookway")
add("D 使用规则", "优惠同享", "不可同享 / 可与全部同享 / 只与部分同享", "美团 抖音 快手", "in", None, "stack")
add("D 使用规则", "退款规则", "随时退 / 过期退 / 有条件退（已裁膜不退）", "全部", "in", None, "refund")
add("D 使用规则", "售卖开始时间", "审核通过立即 / 指定日期", "美团 抖音 快手", "in", None, "sale_from")
add("D 使用规则", "售卖结束时间", "自动延期 / 指定日期", "美团 抖音 快手", "in", None, "sale_to")
add("D 使用规则", "投放渠道", "不限制 / 仅直播间（抖音）；建议不限制", "抖音", "in", None, "channel")
add("D 使用规则", "是否收集客户手机号", "需要 / 不需要；建议不需要，降低下单门槛", "抖音", "in", None, "phone")
add("D 使用规则", "核销方式", "扫码 / 券码 / 报手机号；谁有核销权限写清", "全部", "in", None, "redeem")

# E 履约
add("E 履约与售后", "施工时长（小时）", "玻璃膜 2-4；车衣 24-48；改色 48-72；内饰 24-72", "全部", "in", "0.0", "dur")
add("E 履约与售后", "是否留车过夜", "能 / 不能；能的话写清保管责任", "全部", "in", None, "overnight")
add("E 履约与售后", "改期 / 取消规则", "例：提前 24h 免费改期，最多 2 次；爽约金怎么算", "全部", "in", None, "resched")
add("E 履约与售后", "售后响应", "例：24h 内响应，7 天内免费返工", "全部", "in", None, "after")
add("E 履约与售后", "质保凭证", "电子质保卡 / 纸质卡 + 编号规则", "全部", "in", None, "wcard")
add("E 履约与售后", "交付验收标准", "无水泡/无划伤/无起翘/边角包覆到位；交车当面验", "内部", "in", None, "accept")
add("E 履约与售后", "施工前车况记录", "全车拍照+录像存档，避免事后扯皮", "内部", "in", None, "precheck")

# F 素材
add("F 素材与文案", "商品头图", "美团 16:9 宽高>480px；抖音 ≥2 张；小红书 3:4 竖图；均不可带导流文字", "全部", "in", None, "img_main")
add("F 素材与文案", "环境图", "门头 / 工位 / 休息区，1-3 张，真实拍摄", "全部", "in", None, "img_env")
add("F 素材与文案", "施工过程图", "3-5 张：清洁、裁膜、刮水、收边；无水印无二维码", "全部", "in", None, "img_work")
add("F 素材与文案", "前后对比图", "同角度同光线，3 组以上；这是贴膜店最有效的素材", "全部", "in", None, "img_before")
add("F 素材与文案", "色卡图（改色膜）", "每个颜色一张，≥640×360；最多 5 组", "美团 淘宝", "in", None, "img_color")
add("F 素材与文案", "竖版短视频", "9:16，1080×1920，15-30s；无水印、无二维码、无第三方平台 logo", "抖音 快手 小红书", "in", None, "video")
add("F 素材与文案", "视频封面", "3:4 或 9:16，带大字标题（能看清卖点）", "抖音 快手", "in", None, "cover")
add("F 素材与文案", "小红书笔记标题", "≤20 字。人群+场景+数字，例：天津贴膜别乱选，这3点看完成本省一半", "小红书", "in", None, "xhs_title")
add("F 素材与文案", "小红书笔记正文", "≤1000 字，600 字左右最好；分 3-4 段，每段 2-3 行", "小红书", "in", None, "xhs_body")
add("F 素材与文案", "话题标签", "3-6 个：#天津汽车贴膜 #隐形车衣 #改色膜 #汽车美容", "小红书 抖音 快手", "in", None, "hashtag")
add("F 素材与文案", "地点 POI", "门店 POI 名称，与 02 表一致", "小红书 抖音 快手", "in", None, "poi")

# G 合规
add("G 合规自检", "是否含绝对化用语", "填“无”或列出（最/第一/永久/100%）；有就必须改掉", "全部", "in", None, "c1")
add("G 合规自检", "是否承诺效果/寿命", "填“无”或列出；不承诺“永不发黄”“绝对不留胶”", "全部", "in", None, "c2")
add("G 合规自检", "图片是否含第三方水印/二维码", "填“否”或列出；快手对这条查得最严", "全部", "in", None, "c3")
add("G 合规自检", "宣传价与团购价是否一致", "填“一致”或“不一致（差 XX 元，已改）”", "全部", "in", None, "c4")
add("G 合规自检", "是否引导站外交易", "填“否”或列出；引导加微信会被限流/下架", "全部", "in", None, "c5")

r = 4
rowmap = {}
prev_cat = None
for cat, field, req, plat, kind, fmt, key in F:
    if cat != prev_cat:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
        c = ws.cell(row=r, column=1, value=cat)
        c.font = Font(name=FONT, size=10, bold=True, color="FF1F3864")
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in range(1, 12):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=XL_LIGHT)
            ws.cell(row=r, column=col).border = BORDER
        ws.row_dimensions[r].height = 22
        r += 1
        prev_cat = cat
    rowmap[key] = r
    c1 = ws.cell(row=r, column=1, value=cat)
    c1.font = Font(name=FONT, size=9, color="FF808080")
    c2 = ws.cell(row=r, column=2, value=field)
    c2.font = Font(name=FONT, size=10, bold=True)
    c3 = ws.cell(row=r, column=3, value=req)
    c3.font = Font(name=FONT, size=9, color="FF404040")
    c4 = ws.cell(row=r, column=4, value=plat)
    c4.font = Font(name=FONT, size=9, color="FF404040")
    for i in (1, 2, 3, 4):
        ws.cell(row=r, column=i).alignment = Alignment(
            horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws.cell(row=r, column=i).border = BORDER
    for col in range(5, 12):
        c = ws.cell(row=r, column=col)
        c.border = BORDER
        if kind == "calc":
            c.fill = PatternFill("solid", fgColor=XL_CALC)
        elif col < 11:
            c.fill = PatternFill("solid", fgColor=XL_INPUT)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        if fmt:
            c.number_format = fmt
    ws.row_dimensions[r].height = 30
    r += 1

last = r - 1
# 计算列公式（E..J）
for col in range(5, 11):
    L = get_column_letter(col)
    ws[f"{L}{rowmap['settle']}"] = (
        f'=IF({L}{rowmap["price"]}="","",ROUND({L}{rowmap["price"]}*0.94,2))')
    ws[f"{L}{rowmap['gross']}"] = (
        f'=IF(OR({L}{rowmap["settle"]}="",{L}{rowmap["cost"]}=""),"",ROUND({L}{rowmap["settle"]}-{L}{rowmap["cost"]},2))')
    ws[f"{L}{rowmap['margin']}"] = (
        f'=IF(OR({L}{rowmap["settle"]}="",{L}{rowmap["settle"]}=0,{L}{rowmap["gross"]}=""),"",{L}{rowmap["gross"]}/{L}{rowmap["settle"]})')

ws.freeze_panes = "E4"
ws.auto_filter.ref = f"A3:K{last}"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "3:3"


# ============================================================
# 04_平台必填对照
# ============================================================
ws = wb.create_sheet("04_平台必填对照")
ws.sheet_properties.tabColor = xl_color("#70AD47")
headers = ["字段", "分类", "美团·点评", "抖音来客", "快手", "小红书", "闲鱼", "淘宝", "依据 / 差异说明"]
widths = [22, 14, 11, 11, 11, 11, 11, 11, 52]
title_row(ws, "字段 × 平台 必填对照", 9,
          sub="绿色底＝该平台必填；白底＝选填；“—”＝该平台没有这个字段；橙色＝需在后台实测确认。同一件事各平台叫法不同，看说明列。")
header_row(ws, 3, headers, widths, height=30)

M = [
    # 字段, 分类, 美团, 抖音, 快手, 小红书, 闲鱼, 淘宝, 说明
    ("套餐/商品名称", "基础", "必填", "必填", "必填", "必填", "必填", "必填",
     "美团名由系统按“品牌+分类+质保+材质+特色”自动拼接，不能手改；抖音/快手/淘宝可自定义标题，30 字内。"),
    ("商品标题（独立字段）", "基础", "—", "必填", "必填", "—", "必填", "必填",
     "快手要求含门店品牌简称+套餐名+规格；闲鱼前 15 字最关键；淘宝 30 字/60 字符。小红书标题属笔记，不在商品上。"),
    ("商品类型", "基础", "必填", "必填", "必填", "选填", "选填", "—",
     "团购套餐/代金券/次卡。闲鱼走“发服务”或“发闲置”入口即等于选了类型。"),
    ("平台类目", "基础", "必填", "必填", "必填", "必填", "必填", "必填",
     "填错类目直接影响流量与审核，路径见 05 表。"),
    ("商品头图", "素材", "必填", "必填", "必填", "必填", "必填", "必填",
     "美团 16:9 且宽高>480px；抖音/快手 ≥2 张、禁黑边禁截图；小红书 3:4 竖图；淘宝首图建议白底 800×800 以上。"),
    ("商品视频", "素材", "选填", "选填", "选填", "选填", "选填", "选填",
     "有视频的转化率明显更高，优先拍 15-30s 竖版。"),
    ("图文详情", "素材", "必填", "选填", "必填", "选填", "必填", "必填",
     "美团≤1000 字；快手要求套餐单品明细；闲鱼写在“描述”里；淘宝写在“详情描述”。"),
    ("门市价（原价）", "价格", "必填", "必填", "必填", "必填", "选填", "必填",
     "必须真实可查。虚高再打折是各平台共同红线。"),
    ("团购价 / 售价", "价格", "必填", "必填", "必填", "必填", "必填", "必填",
     "跨平台保持一致。"),
    ("库存", "价格", "选填", "选填", "选填", "选填", "必填", "必填",
     "团购平台可不限；闲鱼/淘宝必须填数量。"),
    ("建议人数 / 车辆数", "价格", "—", "必填", "必填", "—", "—", "—",
     "抖音叫“建议用餐人数”、快手叫“建议人数”，贴膜店填 1 车。"),
    ("品牌", "属性", "必填", "选填", "必填", "选填", "选填", "必填",
     "美团为下拉单选（可新增）；快手要求标题含品牌；淘宝是类目属性必填。"),
    ("型号 / 系列", "属性", "必填", "选填", "选填", "选填", "选填", "必填",
     "美团品牌与型号合并为“品牌及型号”一个字段。"),
    ("材质", "属性", "必填", "选填", "必填", "—", "—", "必填",
     "TPU/PVC 等。美团为下拉可新增。"),
    ("净膜厚度（mil）", "属性", "必填", "—", "—", "—", "—", "选填",
     "只有美团车衣类目有这个字段，且是必填。"),
    ("修复类型", "属性", "必填", "—", "—", "—", "—", "—", "美团车衣专属：自修复/热修复。"),
    ("表面质感", "属性", "必填", "—", "—", "—", "—", "选填", "美团车衣：光面/雾面。"),
    ("耐变黄时间", "属性", "选填", "—", "—", "—", "—", "—", "美团车衣选填。"),
    ("生产工艺", "属性", "必填", "—", "—", "—", "—", "—", "美团改色膜专属：压延/铸造。"),
    ("颜色 / 图案", "属性", "必填", "选填", "选填", "选填", "选填", "必填",
     "改色膜在美团要每个颜色配一张图（最多 5 组）；淘宝走 SKU。"),
    ("质保年限", "属性", "必填", "选填", "选填", "选填", "选填", "必填",
     "美团为下拉（5年/10年，可新增）；没有独立字段的平台就写进标题或详情。"),
    ("质保范围", "属性", "必填", "选填", "选填", "选填", "选填", "选填",
     "写清“保什么、不保什么”，这是差评高发点。"),
    ("适用车型", "属性", "必填", "选填", "选填", "选填", "选填", "必填",
     "美团为下拉：全车型/五座轿车/SUV·MPV/仅SUV/仅MPV。"),
    ("施工流程", "属性", "必填", "选填", "必填", "选填", "选填", "选填",
     "美团是固定多选项；快手要写在套餐详情里。"),
    ("增值服务", "属性", "选填", "选填", "选填", "选填", "选填", "选填", "赠送项，能提转化。"),
    ("图文详情字数", "素材", "≤1000 字", "无硬限", "无硬限", "≤1000 字", "≤2000 字", "无硬限",
     "小红书 1000 字是硬上限；闲鱼描述建议 300-800 字。"),
    ("适用门店", "履约", "必填", "必填", "必填", "必填", "发货地", "必填",
     "前三者都要选可核销门店；闲鱼填所在地；淘宝填服务范围或门店。"),
    ("有效期", "规则", "必填", "必填", "必填", "必填", "选填", "必填",
     "美团可选“购买后 N 天”或“固定区间”；抖音叫“客户可消费日期”。"),
    ("可用时段", "规则", "必填", "必填", "必填", "必填", "—", "选填",
     "美团合并为“不可用日期”；抖音叫“每日消费时段”。"),
    ("预约规则", "规则", "必填", "必填", "必填", "必填", "选填", "选填",
     "车衣/改色必须写“需提前 N 天预约，需留车 N 天”。"),
    ("优惠同享", "规则", "必填", "选填", "选填", "选填", "—", "—", "美团必填三选一。"),
    ("限购", "规则", "必填", "必填", "必填", "必填", "—", "选填", "美团/抖音/快手都必填。"),
    ("退款规则", "规则", "平台默认", "必填", "必填", "必填", "平台默认", "必填",
     "注意：已裁膜/已施工通常不可退，必须提前写清。"),
    ("售卖起止时间", "规则", "必填", "必填", "必填", "选填", "—", "必填", "美团/抖音建议“自动延期”。"),
    ("核销方式", "履约", "平台券", "必填", "必填", "必填", "—", "选填",
     "扫码/券码/报手机号；抖音/快手需先配核销员账号。"),
    ("达人佣金", "推广", "选填", "必填*", "选填", "—", "—", "—",
     "开了达人带货才需填；抖音团购达人佣金常见 10%-20%。"),
    ("投放渠道", "推广", "—", "必填", "—", "—", "—", "—", "抖音：不限制 / 仅直播间。"),
    ("是否收集客户手机号", "规则", "—", "必填", "—", "—", "—", "—", "抖音建议选“不需要”。"),
    ("发货 / 交付方式", "履约", "—", "—", "—", "—", "必填", "必填",
     "闲鱼选自提/邮寄；淘宝服务类选“无需物流”或设置服务范围。"),
    ("服务承诺", "规则", "—", "—", "—", "—", "选填", "必填",
     "淘宝有“7 天无理由/破损包退”等；服务类通常不适用 7 天无理由，按类目提示选。"),
    ("运费模板", "履约", "—", "—", "—", "—", "选填", "必填", "淘宝实物膜才需要；纯服务类用不到。"),
    ("成色 / 新旧", "属性", "—", "—", "—", "—", "必填", "—", "闲鱼服务类选“全新”或“不适用”。"),
    ("所在地 / 可服务区域", "履约", "—", "—", "—", "—", "必填", "—", "闲鱼同城流量靠这个；建议写“天津市西青区”。"),
    ("鱼塘 / 标签", "推广", "—", "—", "—", "—", "选填", "—", "同步到本地鱼塘能显著增加曝光。"),
    ("笔记标题（≤20 字）", "内容", "—", "—", "—", "必填", "—", "—", "小红书专属，超 20 字发不出去。"),
    ("笔记正文（≤1000 字）", "内容", "—", "—", "—", "必填", "—", "—", "小红书专属，图多字少更吃香。"),
    ("话题标签", "内容", "—", "选填", "选填", "必填", "选填", "—", "小红书 3-5 个；抖音/快手 3-6 个。"),
    ("地点 POI 挂载", "内容", "—", "必填", "必填", "必填", "—", "—",
     "门店 POI 挂载是本地流量的入口，抖音/快手/小红书都要。"),
    ("商品卡 / 团购挂载", "内容", "—", "必填", "必填", "必填", "—", "—",
     "视频/笔记里挂团购组件，才能“边看边买”。"),
    ("企业资质认证", "主体", "必填", "必填", "必填", "必填", "实名", "必填",
     "抖音蓝V 600 元/年；小红书企业专业号 600 元/年；美团/点评入驻免费按交易抽佣。"),
    ("门店 POI 认领", "主体", "必填", "必填", "必填", "必填", "—", "—",
     "POI 名称/地址/门头照三者必须对得上，否则反复驳回。"),
]

r = 4
prev_cat = None
for row in M:
    if row[1] != prev_cat:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        c = ws.cell(row=r, column=1, value=row[1])
        c.font = Font(name=FONT, size=10, bold=True, color="FF1F3864")
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in range(1, 10):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=XL_LIGHT)
            ws.cell(row=r, column=col).border = BORDER
        ws.row_dimensions[r].height = 22
        r += 1
        prev_cat = row[1]
    for i, v in enumerate(row, start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.border = BORDER
        if i == 1:
            c.font = Font(name=FONT, size=10, bold=True)
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        elif i == 2:
            c.font = Font(name=FONT, size=9, color="FF808080")
            c.alignment = Alignment(horizontal="center", vertical="center")
        elif i == 9:
            c.font = Font(name=FONT, size=9, color="FF404040")
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        else:
            c.font = Font(name=FONT, size=9, bold=False)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 30
    r += 1

last = r - 1
rng = f"C4:H{last}"
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"必填"'],
                                              fill=PatternFill("solid", fgColor=XL_GREEN_F),
                                              font=Font(name=FONT, size=9, bold=True, color=XL_GREEN_T)))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"必填*"'],
                                              fill=PatternFill("solid", fgColor=XL_GREEN_F),
                                              font=Font(name=FONT, size=9, color=XL_GREEN_T)))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"待核实"'],
                                              fill=PatternFill("solid", fgColor=XL_WARN),
                                              font=Font(name=FONT, size=9, color=XL_RED_T)))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"—"'],
                                              font=Font(name=FONT, size=9, color=XL_GRAY_T)))
ws.freeze_panes = "C4"
ws.auto_filter.ref = f"A3:I{last}"


# ============================================================
# 05_平台字段明细
# ============================================================
ws = wb.create_sheet("05_平台字段明细")
ws.sheet_properties.tabColor = xl_color("#FFC000")
widths = [4, 24, 10, 16, 46, 40]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
title_row(ws, "各平台后台字段明细（上架时对着填）", 6,
          sub="按平台分块。每块就是该平台创建商品时的表单顺序，尽量按后台顺序排，减少来回找。")

blocks = [
    ("美团 / 大众点评 · 经营宝（两平台共用同一后台）", [
        ("类目路径", "必填", "下拉", "爱车 > 美容 > 隐形车衣 或 车身改色", "类目决定能填哪些属性，先选对类目再填内容"),
        ("团购名称", "必填", "系统拼接", "品牌+分类+质保年限+材质+特色，最多 30 字，不能手改", "所以“品牌”和“质保年限”选错，名字就错了"),
        ("团购券格式", "必填", "单选", "点评券 / 三方券", "自营核销一般选点评券"),
        ("团购售卖类型", "必填", "单选", "单次卡 / 多次卡", "贴膜按次卖选单次卡"),
        ("团购头图", "必填", "图片", "建议 16:9，宽高均 >480px，5M 内，不可加导向性文字", "带“加微信/立减”等字样会被驳回"),
        ("团购视频", "选填", "视频", "服务讲解 / 施工案例 / 养车科普 / 店内带看，需审核", ""),
        ("门市价", "必填", "金额", "原价", "必须真实"),
        ("团购价", "必填", "金额", "实际售卖价", "平台基础服务费 6%"),
        ("库存", "选填", "整数", "默认不限", "建议按产能设，避免接不住"),
        ("适用位置", "必填", "单选", "全车 / 可自定义（最多 1 项）", ""),
        ("净膜厚度", "必填*", "数字+mil", "车衣专属，如 7.5", "改色膜无此字段"),
        ("修复类型", "必填*", "单选", "自修复 / 热修复 / 可自定义", "车衣专属"),
        ("服务流程", "必填", "多选", "整车外表清洁 / 外观车况检查 / 专车专用裁膜 / 施工贴膜 / 可自定义", ""),
        ("质保范围", "必填", "文本", "手动填写覆盖内容，如 脱胶、起泡", ""),
        ("增值服务", "选填", "多选", "赠送车衣险 / 赠送 3 次免费修补 / 可自定义", ""),
        ("表面质感", "必填*", "单选", "光面 / 雾面 / 可自定义", "车衣专属"),
        ("耐变黄时间", "选填", "单选", "5 年 / 10 年 / 可自定义", ""),
        ("生产厂商", "选填", "文本", "如 卡皮诺", ""),
        ("份数", "必填", "整数+份", "该服务项目数量", ""),
        ("改色膜类型", "必填*", "单选", "哑光膜 / 电镀膜 / 可自定义", "改色膜专属"),
        ("改色膜材质", "必填*", "单选", "PVC / 可自定义", "改色膜专属"),
        ("图案风格", "必填*", "单选", "卡通动漫 / 可自定义", "改色膜专属"),
        ("生产工艺", "必填*", "单选", "压延 / 铸造 / 可自定义", "改色膜专属"),
        ("车衣品牌及型号", "必填", "单选（可新增）", "如 卡皮诺 / 3M / VPTINT", "既是名字来源也是属性"),
        ("车衣材质", "必填*", "单选", "TPU / PVC / 可新增", "车衣专属"),
        ("车膜颜色", "必填*", "图片+文本", "每色一张图，≥640×360，16:9，5M 内，最多 5 组", "改色膜专属"),
        ("质保年限", "必填", "单选", "5 年质保 / 10 年质保 / 可新增", ""),
        ("适用车型", "必填", "单选", "全车型 / 五座轿车 / SUV/MPV / 仅 SUV / 仅 MPV", ""),
        ("图文详情", "必填", "图文", "图片 16:9、5M 内；文字最多 1000 字", ""),
        ("补充信息", "选填", "文本", "对套餐内容的补充说明", ""),
        ("适用门店", "必填", "门店选择", "可核销门店列表", "多店要全选上，否则用户买不了"),
        ("团购使用有效期", "必填", "单选+数值", "指定天数内有效 / 指定时间段内有效", ""),
        ("不可用日期", "必填", "单选", "全部时间可用 / 限制时间", ""),
        ("是否需要预约", "必填", "单选", "不需要 / 每天规则相同 / 节假日与周末规则不同", ""),
        ("优惠同享信息", "必填", "单选", "不可同享 / 可全部同享 / 只与部分同享", ""),
        ("单人购买量限制", "必填", "单选", "不限制 / 设上限", ""),
        ("售卖开始时间", "必填", "单选", "审核通过立即售卖 / 指定时间", ""),
        ("售卖结束时间", "必填", "单选", "自动延期保持售卖 / 指定结束时间", "建议自动延期"),
    ]),
    ("抖音来客（团购套餐）", [
        ("商品品类", "必填", "按执照经营范围自动匹配", "如 汽车服务 > 汽车美容", "品类选错，后面属性全错"),
        ("商品类型", "必填", "单选", "团购套餐 / 代金券 / 次卡", "贴膜选团购套餐"),
        ("适用门店", "必填", "门店选择", "已认领授权的门店", "没认领 POI 在这里会选不到"),
        ("套餐名称", "必填", "文本", "建议【品牌】+项目+规格+质保，30 字内", ""),
        ("建议人数", "必填", "数字", "贴膜店填 1 车（餐饮才填人数）", ""),
        ("商品搭配", "必填", "结构化多行", "商品组名 + 单品名称 + 店内单价；可多组", "对应“套餐包含什么”，这是审核重点"),
        ("顾客实际支付", "必填", "金额", "团购价", ""),
        ("商品图", "必填", "图片", "至少 2 张套餐图；建议 1:1 或 3:4，无水印", ""),
        ("环境图", "选填", "图片", "门店/车间实拍", ""),
        ("库存", "选填", "整数", "限制 / 不限", ""),
        ("投放渠道", "必填", "单选", "不限制 / 仅直播间", "建议不限制"),
        ("商品售卖日期", "必填", "日期区间", "默认一年，可自动延期", ""),
        ("客户可消费日期", "必填", "单选+数值", "购买后 N 天内可用（建议 90 天）", ""),
        ("每日消费时段", "必填", "时间区间", "默认全天，建议按营业时间设", ""),
        ("限购规则", "必填", "单选+数值", "每人限购 N 份", ""),
        ("预约规则", "必填", "单选", "无需预约 / 提前 N 小时 / 提前 N 天", "车衣务必设提前 1-3 天"),
        ("服务方式", "必填", "单选", "到店服务（餐饮为堂食/外带）", ""),
        ("使用规则", "必填", "多行文本", "可写多条：不可用日期、加价规则、质保说明", "写得越清，退款越少"),
        ("其他说明", "选填", "文本", "补充信息", ""),
        ("顾客信息设置", "必填", "单选", "需要 / 不需要客户手机号", "建议不需要"),
        ("达人佣金", "选填", "百分比", "开了达人带货才需填，常见 10%-20%", ""),
        ("核销员 / 权限", "必填", "配置", "给店员单独开核销权限，别共用主账号", ""),
        ("收款账户", "必填", "绑定", "不绑定无法创建商品", ""),
    ]),
    ("快手本地生活", [
        ("门店认领 / 新建", "必填", "结构化", "门店名称、类目、地址、电话、举证图片", "地址要与执照/地图一致"),
        ("主体资质", "必填", "上传", "营业执照、行业资质（按主营）、法人身份信息", ""),
        ("商品类目", "必填", "下拉", "按实际属性选，别为了流量乱选", "类目错→流量偏、达人不敢接"),
        ("商品标题", "必填", "文本", "完整含门店品牌（简称）+套餐名+规格；多店通用要注明“N 店通用”", "不得出现与商品无关的信息"),
        ("商品图片", "必填", "图片", "套餐图、菜品图（此处为项目图）、环境图；高清原图，禁黑边、禁截图、禁第三方水印", "快手查得最严，含“大众点评”“小红书”字样直接驳"),
        ("商品搭配 / 套餐详情", "必填", "结构化", "每个单品明细：名称、数量、门店价；套餐整体门市价/团购价", ""),
        ("商品价格", "必填", "金额", "门市价 + 团购价", ""),
        ("库存", "选填", "整数", "数量或不限", ""),
        ("有效期", "必填", "单选+数值", "购买后 N 天 / 固定区间", ""),
        ("可用时段 / 不可用日期", "必填", "时间", "营业时间内；节假日是否可用", ""),
        ("预约规则", "必填", "单选", "是否需预约、提前多久", ""),
        ("限购", "必填", "单选", "每人限购数量", ""),
        ("使用规则 / 注意事项", "必填", "文本", "写清加价、留车、质保边界", ""),
        ("退款规则", "必填", "单选", "随时退 / 过期退 / 有条件退", ""),
        ("核销方式", "必填", "配置", "扫码 / 券码；需配核销账号", ""),
        ("商家产品编码", "选填", "文本", "接入第三方系统（如智慧门店）时填对应 ID", "纯手工经营不用管"),
        ("达人佣金", "选填", "百分比", "开达人带货时填", ""),
    ]),
    ("小红书（企业专业号 + 本地生活团购 + 笔记）", [
        ("企业专业号认证", "必填", "认证", "600 元/年（审核费，失败不退）", "个人号开不了本地生活"),
        ("POI 门店认领", "必填", "结构化", "门店名称、地址、电话、营业时间、门头照", "门头照用网图会被驳回"),
        ("本地生活类目", "待核实", "下拉", "汽车服务 / 洗车养护 归属“生活服务”；能否自助开通需后台确认", "部分类目需服务商报白，见 01 表说明"),
        ("团购商品名称", "必填", "文本", "项目+规格+质保，20-30 字", ""),
        ("团购价 / 门市价", "必填", "金额", "", ""),
        ("商品主图", "必填", "图片", "3:4 竖图，建议 ≥1080×1440，无水印", "竖图是小红书的基本盘"),
        ("团购详情", "必填", "结构化/文本", "包含的服务项、数量、使用门店", ""),
        ("使用规则", "必填", "文本", "有效期、可用时段、预约、退款、加价", ""),
        ("适用门店", "必填", "门店选择", "认领过的 POI", ""),
        ("库存 / 限购", "选填", "整数", "", ""),
        ("笔记标题", "必填", "文本", "≤20 字（含标点/emoji）", "超了存不下来"),
        ("笔记正文", "选填", "文本", "≤1000 字；600 字左右、两行一空行最好读", ""),
        ("笔记图片", "必填", "图片", "1-18 张，3:4 竖图为主，首图为封面（决定点开率）", ""),
        ("话题标签", "选填", "多选", "3-5 个，一大流量+几个精准词", ""),
        ("地点 POI", "必填", "选择", "门店 POI", "同城流量入口"),
        ("商品卡 / 店铺卡", "必填", "挂载", "发布时挂团购组件", ""),
        ("私信 / 群聊组件", "选填", "组件", "承接咨询，别在正文留微信", ""),
    ]),
    ("闲鱼（发服务）", [
        ("发布入口", "必填", "入口选择", "发服务（技能/本地服务）或 发闲置", "选错入口字段对不上"),
        ("账号实名", "必填", "认证", "实名 + 支付宝绑定 + 芝麻信用授权", "未实名不能发"),
        ("类目", "必填", "下拉", "服务 > 汽车服务 > 汽车美容/贴膜（以实际下拉为准）", "类目树随版本变动"),
        ("标题", "必填", "文本", "≤30 字；前 15 字放核心词：天津汽车贴膜+项目+车型", "算法抓前 15 字"),
        ("描述", "必填", "文本", "≤2000 字。分段写：包含什么/品牌型号/施工流程/质保/预约/地址/加价规则", ""),
        ("图片", "必填", "图片", "1-9 张，首图无水印、主体清晰；建议 3 张以上", ""),
        ("视频", "选填", "视频", "15-60s，有视频曝光更好", ""),
        ("价格", "必填", "金额", "一口价或可议价，支持两位小数", ""),
        ("原价", "选填", "金额", "划线价", ""),
        ("库存数量", "必填", "整数", "可接单数", ""),
        ("交付方式", "必填", "单选", "自提（到店）/ 邮寄 / 无需物流", "服务选到店自提最合适"),
        ("所在地", "必填", "地区", "天津市西青区", "决定同城曝光"),
        ("成色 / 新旧", "必填", "单选", "全新 / 不适用", "服务类选全新"),
        ("标签", "选填", "多选", "如 汽车美容、贴膜、隐形车衣", "影响推荐权重"),
        ("鱼塘同步", "选填", "选择", "同步到本地汽车/天津鱼塘", "本地曝光提升明显"),
        ("服务保障", "选填", "勾选", "如实描述、极速退款等", ""),
    ]),
    ("淘宝（服务类目 / 实物膜）", [
        ("店铺类型", "必填", "单选", "企业店 / 个人店（服务类目建议企业店）", ""),
        ("类目", "必填", "下拉", "服务 > 汽车服务 > 汽车装饰与美容 / 安装服务；卖实物膜则走汽车用品", "两类打法不同：卖服务 vs 卖膜料"),
        ("资质", "必填", "上传", "营业执照（经营范围含汽车美容/维修）、行业许可（如有）、技师资格（部分类目）", "经营范围不含会被驳"),
        ("宝贝标题", "必填", "文本", "30 字 / 60 字符，项目词+品牌+规格+质保+地域", ""),
        ("副标题", "选填", "文本", "", ""),
        ("类目属性", "必填", "下拉/文本", "品牌、型号、材质、颜色、适用车型、质保年限、厚度", "随类目不同，以系统标红必填为准"),
        ("SKU 规格", "必填", "多行", "按车型/颜色/位置分 SKU，各自价格与库存", "贴膜店最常见：轿车/SUV 两档 SKU"),
        ("主图", "必填", "图片", "5 张，首图建议白底 800×800 以上；无牛皮癣文字", ""),
        ("主图视频", "选填", "视频", "9-16s", ""),
        ("详情描述", "必填", "图文", "建议 10-15 屏：包含项、施工流程、质保、门店实拍、常见问题", ""),
        ("价格 / 区间价", "必填", "金额", "一口价或按 SKU 区间", ""),
        ("库存", "必填", "整数", "按 SKU 分别设", ""),
        ("运费 / 服务范围", "必填", "模板", "实物膜设运费模板；纯服务类设服务范围/无需物流", ""),
        ("服务承诺", "必填", "勾选", "7 天无理由（服务类多不适用）/ 破损包退 / 安装保障", "选错承诺会带来纠纷"),
        ("上架时间", "必填", "日期时间", "立即或定时", ""),
    ]),
]

r = 3
for bname, rows_ in blocks:
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c = ws.cell(row=r, column=1, value=bname)
    c.font = Font(name=FONT, size=11, bold=True, color=XL_WHITE)
    c.fill = PatternFill("solid", fgColor=xl_color("#2F5597"))
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(1, 7):
        ws.cell(row=r, column=col).border = BORDER
        ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=xl_color("#2F5597"))
    ws.row_dimensions[r].height = 26
    r += 1
    # 表头
    for i, h in enumerate(["#", "字段", "必填", "类型", "可选值 / 规格限制", "备注 / 易驳回点"], start=1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = Font(name=FONT, size=10, bold=True, color="FF1F3864")
        c.fill = PatternFill("solid", fgColor=XL_LIGHT)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER
    ws.row_dimensions[r].height = 22
    r += 1
    for idx, (field, must, typ, opts, note) in enumerate(rows_, start=1):
        vals = [idx, field, must, typ, opts, note]
        for i, v in enumerate(vals, start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.border = BORDER
            if i == 1:
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = Font(name=FONT, size=9, color="FF808080")
            elif i == 2:
                c.font = Font(name=FONT, size=10, bold=True)
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            elif i == 3:
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                bold = v.startswith("必填")
                c.font = Font(name=FONT, size=9, bold=bold,
                              color=XL_GREEN_T if bold else (XL_RED_T if v == "待核实" else "FF404040"))
                if v == "待核实":
                    c.fill = PatternFill("solid", fgColor=XL_WARN)
            elif i == 4:
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.font = Font(name=FONT, size=9, color="FF404040")
            elif i == 5:
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
                c.font = Font(name=FONT, size=9)
            else:
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
                c.font = Font(name=FONT, size=9, color="FF595959")
        ws.row_dimensions[r].height = 32
        r += 1
    r += 1  # 块间空一行

ws.sheet_view.showGridLines = False


# ============================================================
# 06_素材清单
# ============================================================
ws = wb.create_sheet("06_素材清单")
ws.sheet_properties.tabColor = xl_color("#7030A0")
headers = ["编号", "素材类型", "用途平台", "规格要求", "需要数量", "状态", "文件名 / 存放位置", "备注"]
widths = [8, 24, 22, 46, 10, 12, 30, 26]
title_row(ws, "素材清单（拍一次，六平台复用）", 8,
          sub="“状态”列是下拉，选完自动变色。缺的素材先标“待拍”，上架前必须清零——没素材的平台上了也没转化。")
header_row(ws, 3, headers, widths)

mats = [
    ("M01", "门头照", "全部（门店认证）", "横版，招牌完整清晰，无遮挡无贴纸，≥1080px 宽", 1, "待拍", "", "所有平台都要，且不能与营业执照照片混用"),
    ("M02", "店内环境照", "全部", "工位、车间全景，横版 ≥1080px", 2, "待拍", "", ""),
    ("M03", "休息区 / 接待区", "美团 点评 小红书", "光线充足，横版", 1, "待拍", "", ""),
    ("M04", "营业执照", "全部", "四角完整、不反光、字迹清晰", 1, "待拍", "", ""),
    ("M05", "法人身份证正反面", "全部", "清晰、不遮挡", 2, "待拍", "", ""),
    ("M06", "法人手持身份证照", "部分平台", "按平台示例拍", 1, "待拍", "", ""),
    ("M07", "品牌授权书", "美团 抖音 淘宝", "雷朋/3M/科为等，拍照或扫描", 0, "不需要", "", "有就传，能提审核通过率"),
    ("M08", "膜料包装 / 卷材图", "全部", "能看清品牌 logo 与型号", 3, "待拍", "", "证明用的是正品膜"),
    ("M09", "施工过程图", "全部", "清洁、裁膜、刮水、收边各 1 张，横版", 4, "待拍", "", "无水印无二维码"),
    ("M10", "前后对比图·玻璃膜", "全部", "同角度同光线，车内看出去的效果", 3, "待拍", "", "最有效的素材类型"),
    ("M11", "前后对比图·隐形车衣", "全部", "提亮/防刮细节特写 + 全车", 3, "待拍", "", ""),
    ("M12", "前后对比图·改色膜", "全部", "全车正视 + 侧面，同机位", 3, "待拍", "", ""),
    ("M13", "色卡图（改色膜）", "美团 淘宝 小红书", "每个颜色一张 ≥640×360，16:9", 5, "待拍", "", "美团最多 5 组"),
    ("M14", "内饰改装完成图", "全部", "座椅、门板、顶棚细节", 4, "待拍", "", ""),
    ("M15", "音响 / 隔音施工图", "全部", "拆装、隔音棉铺贴、完工", 3, "待拍", "", ""),
    ("M16", "短视频·门店带看", "抖音 快手 小红书", "9:16，1080×1920，15-30s，无水印", 1, "待拍", "", "置顶用"),
    ("M17", "短视频·施工过程", "抖音 快手 小红书", "9:16，15-30s，有节奏感", 3, "待拍", "", ""),
    ("M18", "短视频·前后对比", "抖音 快手 小红书", "9:16，15-30s，前 3 秒出对比", 3, "待拍", "", ""),
    ("M19", "视频封面模板", "抖音 快手", "3:4 或 9:16，大字标题", 1, "待拍", "", "统一风格，便于识别"),
    ("M20", "价目表图", "小红书 私信", "竖版，清晰可读", 1, "待拍", "", "别放平台内当商品图（易判导流）"),
    ("M21", "质保卡样式", "全部", "电子版/实物拍照", 1, "待拍", "", ""),
    ("M22", "客户好评截图", "小红书 美团", "需客户授权，打码手机号", 5, "待拍", "", "未经授权别用"),
]
r = 4
for row in mats:
    for i, v in enumerate(row, start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.border = BORDER
        c.font = Font(name=FONT, size=10, bold=(i == 2))
        if i in (1, 5):
            c.alignment = Alignment(horizontal="center", vertical="center")
        elif i == 6:
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.fill = PatternFill("solid", fgColor=XL_INPUT)
        elif i == 7:
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            c.fill = PatternFill("solid", fgColor=XL_INPUT)
        elif i == 8:
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            c.font = Font(name=FONT, size=9, color="FF595959")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    ws.row_dimensions[r].height = 28
    r += 1
last = r - 1

dv = DataValidation(type="list", formula1='"已有,待拍,待修,不需要"', allow_blank=True,
                    showErrorMessage=True, errorTitle="只能选", error="请从下拉里选：已有/待拍/待修/不需要")
ws.add_data_validation(dv)
dv.add(f"F4:F{last}")

ws.conditional_formatting.add(f"F4:F{last}", CellIsRule(operator="equal", formula=['"已有"'],
                                                        fill=PatternFill("solid", fgColor=XL_GREEN_F),
                                                        font=Font(name=FONT, size=10, color=XL_GREEN_T)))
ws.conditional_formatting.add(f"F4:F{last}", CellIsRule(operator="equal", formula=['"待拍"'],
                                                        fill=PatternFill("solid", fgColor=XL_RED_F),
                                                        font=Font(name=FONT, size=10, color=XL_RED_T)))
ws.conditional_formatting.add(f"F4:F{last}", CellIsRule(operator="equal", formula=['"待修"'],
                                                        fill=PatternFill("solid", fgColor=XL_YEL_F),
                                                        font=Font(name=FONT, size=10, color=XL_YEL_T)))
ws.conditional_formatting.add(f"F4:F{last}", CellIsRule(operator="equal", formula=['"不需要"'],
                                                        font=Font(name=FONT, size=10, color=XL_GRAY_T)))
ws.freeze_panes = "A4"
ws.auto_filter.ref = f"A3:H{last}"


# ============================================================
# 07_合规与话术
# ============================================================
ws = wb.create_sheet("07_合规与话术")
ws.sheet_properties.tabColor = xl_color("#C55A11")
widths = [16, 40, 50]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
title_row(ws, "合规红线与话术模板", 3,
          sub="贴膜店的文案有 80% 是固定的，把这页的东西背下来，就不会因为一句话被驳回或被差评。")

r = 3
def block_title(text):
    global r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1, value=text)
    c.font = Font(name=FONT, size=11, bold=True, color=XL_WHITE)
    c.fill = PatternFill("solid", fgColor=xl_color("#2F5597"))
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(1, 4):
        ws.cell(row=r, column=col).border = BORDER
        ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=xl_color("#2F5597"))
    ws.row_dimensions[r].height = 24
    r += 1


def block_header(cols):
    global r
    for i, h in enumerate(cols, start=1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = Font(name=FONT, size=10, bold=True, color="FF1F3864")
        c.fill = PatternFill("solid", fgColor=XL_LIGHT)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER
    ws.row_dimensions[r].height = 22
    r += 1


def block_rows(rows_, h=30):
    global r
    for row in rows_:
        for i, v in enumerate(row, start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.border = BORDER
            c.font = Font(name=FONT, size=9 if i > 1 else 10, bold=(i == 1))
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[r].height = h
        r += 1


block_title("一、禁写清单（写了就驳回，或者会被截图挂）")
block_header(["类型", "不要写", "改成这样写"])
block_rows([
    ("绝对化用语", "最好 / 第一 / 顶级 / 唯一 / 永久 / 永不 / 100%", "“5 年质保”“行业常用工艺”——用可验证的事实代替形容词"),
    ("效果承诺", "保证不掉色 / 永不发黄 / 绝对不留胶 / 保证不褪色", "“正常使用下，5 年内出现起泡、脱胶可免费返工”"),
    ("虚假对比", "同行用的都是劣质膜 / 别家都是假货", "“本店使用 XX 品牌授权膜料，可现场验货”"),
    ("价格误导", "原价 9980 现价 1999（原价虚高）", "门市价与团购价都真实，折扣幅度别超过 5 折"),
    ("安全功效", "防爆膜=玻璃不会碎 / 贴了不会出事故", "“玻璃碎裂时减少碎片飞溅”"),
    ("站外导流", "图片/文案里留微信、二维码、手机号", "平台内私信沟通；电话只在门店信息字段里填"),
    ("医疗化表述", "除甲醛 / 杀菌 / 净化空气（膜不具备）", "“隔热、隔紫外线、提升隐私性”"),
], h=34)

r += 1
block_title("二、必须写清的 7 件事（写清了，退款和差评少一半）")
block_header(["#", "要写清的", "写法示例"])
block_rows([
    ("1", "有效期", "购买后 90 天内有效，逾期未使用可申请退款"),
    ("2", "可用时段", "周一至周日 09:00-18:00；节假日正常营业（春节前 3 天停接）"),
    ("3", "是否需预约", "需提前 1 天电话预约；隐形车衣/改色需提前 3 天并预留 1-2 天施工时间"),
    ("4", "加价规则", "SUV 加 200 元，MPV 加 300 元，超跑/改装车另议；已含施工费"),
    ("5", "退款边界", "未施工可随时退；已裁膜或已开始施工不可退"),
    ("6", "质保边界", "保起泡、脱胶、开裂、明显褪色；不含人为剐蹭、事故、鸟粪腐蚀未及时清理"),
    ("7", "留车与保管", "车衣/改色需留车 1-3 天，期间车辆停放在店内车间并有监控；贵重物品请带走"),
], h=32)

r += 1
block_title("三、标题与文案公式")
block_header(["用途", "公式", "示例"])
block_rows([
    ("抖音/快手标题", "【品牌】+项目+规格+车型/质保", "【卡皮诺】隐形车衣全车 TPU 7.5mil 5年质保 天津西青"),
    ("小红书标题", "人群/场景 + 反常识或数字 + 地点", "天津贴膜别乱选｜这3个参数看不懂，等于白花钱"),
    ("闲鱼标题", "地域+项目+车型+卖点（前15字最关键）", "天津汽车贴膜 隐形车衣全车 TPU 5年质保 到店施工"),
    ("淘宝标题", "项目词+品牌+规格+质保+地域（30字内）", "汽车贴膜全车隐形车衣TPU防刮自修复5年质保天津到店"),
    ("一句话卖点", "参数 + 结果，不超过 15 字", "7.5mil 厚度 · 防刮提亮 · 5 年质保"),
    ("详情第一段", "先说包含什么，再说谁适合", "本套餐含：全车漆面精洗 1 次 + 专车裁膜 + 全车贴覆 + 边角包覆 + 5 年质保。适合新车、想保护原厂漆的车主。"),
], h=34)

r += 1
block_title("四、私信首响话术（30 秒内发出，决定成不成单）")
block_header(["场景", "话术", "要点"])
block_rows([
    ("问价（玻璃膜）",
     "您好，玻璃膜我们分三档：入门 8xx、主流 1xxx（质保 5 年）、高隔热 2xxx。方便说下车型和想要的效果吗（隔热/隐私/防爆）？我给您报准确价。",
     "先给区间，再要车型——不要一上来报最低价"),
    ("问价（车衣/改色）",
     "车衣看车型和膜料，轿车 2xxx 起，SUV 上浮 200-300。您是什么车？我发几组同车型的完工图给您看看效果。",
     "用同车型案例代替报价，转化更高"),
    ("问质保",
     "质保 5 年，保起泡、脱胶、开裂和明显褪色，出问题免费返工；人为剐蹭和事故不在范围内。质保卡会写编号，凭卡和车牌就能查。",
     "把“不保什么”也说清楚，反而更可信"),
    ("问能不能便宜",
     "线上这个价是含施工的实价，能送的是：免费旧膜撕除 + 一次内饰清洁。要不您先约个时间到店看膜？不满意不施工。",
     "少降价，多送东西；把决策推到到店"),
    ("客户嫌远 / 犹豫",
     "我们在西青大寺，津港公路边上，导航搜“津豪汽车贴膜改装”就到，门口能停车。您如果时间不方便，可以先加个预约时间，到店不满意随时可退。",
     "给明确导航 + 可退承诺，降低决策成本"),
], h=54)

r += 1
block_title("五、贴膜店专属注意（这几条最容易翻车）")
block_header(["#", "注意点", "怎么处理"])
block_rows([
    ("1", "原车漆/旧膜状况", "施工前全车拍照录像，旧膜老化、漆面有伤的当场和客户确认并记录，避免事后扯皮"),
    ("2", "拆装风险", "内饰改装、音响升级涉及拆装卡扣，事先说明“可能有轻微使用痕迹”，并写入使用规则"),
    ("3", "施工周期", "车衣/改色必须写清留车天数，且说明“以车况为准可能顺延”，别承诺“今天一定好”"),
    ("4", "膜料真伪", "准备授权书与膜料包装照，客户问就发；这比任何文案都有说服力"),
    ("5", "前挡透光率", "前挡按法规需 ≥70%，别推深色前挡膜，也别宣传“全车最深色”"),
    ("6", "雨季/低温", "低温施工需在车间内，写清“气温低于 5℃ 建议改期”，减少返工争议"),
], h=34)

ws.sheet_view.showGridLines = False


# ============================================================
# 08_发布进度
# ============================================================
ws = wb.create_sheet("08_发布进度")
ws.sheet_properties.tabColor = xl_color("#00B050")
headers = ["平台", "账号 / 门店", "套餐", "上架状态", "提交日期", "审核结果",
           "驳回原因 / 待补材料", "商品链接", "负责人", "备注"]
widths = [16, 20, 18, 12, 12, 12, 40, 26, 10, 22]

# 锚点：第1行标题，第2行概览，第3行表头，第4~24行数据（21行）
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
c = ws.cell(row=1, column=1, value="发布进度追踪（自动统计）")
c.font = Font(name=FONT, size=15, bold=True, color=XL_WHITE)
c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[1].height = 34

DATA_START, DATA_END = 4, 24
kpi = [("总项数", f"=COUNTA(A{DATA_START}:A{DATA_END})"),
       ("已上架", f'=COUNTIF(D{DATA_START}:D{DATA_END},"已上架")'),
       ("审核中", f'=COUNTIF(D{DATA_START}:D{DATA_END},"审核中")'),
       ("被驳回", f'=COUNTIF(D{DATA_START}:D{DATA_END},"被驳回")'),
       ("完成率", '=IF(OR(B2="",B2=0),"",D2/B2)')]
col = 1
for label, formula in kpi:
    c = ws.cell(row=2, column=col, value=label)
    c.font = Font(name=FONT, size=9, color="FF595959")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c2 = ws.cell(row=2, column=col + 1, value=formula)
    c2.font = Font(name=FONT, size=12, bold=True, color="FF1F3864")
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.fill = PatternFill("solid", fgColor=XL_CALC)
    c2.border = BORDER
    if label == "完成率":
        c2.number_format = "0.0%"
    col += 2
ws.row_dimensions[2].height = 28

header_row(ws, 3, headers, widths, height=28)

plats = ["美团", "大众点评", "抖音来客", "快手本地生活", "小红书", "闲鱼", "淘宝"]
pkgs = ["玻璃隔热膜", "隐形车衣（TPU）", "车身改色膜"]
r = DATA_START
for p in plats:
    for k in pkgs:
        ws.cell(row=r, column=1, value=p).font = Font(name=FONT, size=10, bold=True)
        ws.cell(row=r, column=2, value="津豪汽车贴膜改装").font = Font(name=FONT, size=10)
        ws.cell(row=r, column=3, value=k).font = Font(name=FONT, size=10)
        c = ws.cell(row=r, column=4, value="未开始")
        c.fill = PatternFill("solid", fgColor=XL_INPUT)
        c.font = Font(name=FONT, size=10)
        ws.cell(row=r, column=5).number_format = "YYYY-MM-DD"
        for col_i in range(1, 11):
            cell = ws.cell(row=r, column=col_i)
            cell.border = BORDER
            if col_i in (1, 3):
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            elif col_i == 7:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 24
        r += 1

dv2 = DataValidation(type="list", formula1='"未开始,准备素材,待提交,审核中,已上架,被驳回,暂不上架"',
                     allow_blank=True, showErrorMessage=True, errorTitle="只能选",
                     error="请从下拉里选")
ws.add_data_validation(dv2)
dv2.add(f"D{DATA_START}:D{DATA_END}")

ws.conditional_formatting.add(f"D{DATA_START}:D{DATA_END}", CellIsRule(
    operator="equal", formula=['"已上架"'], fill=PatternFill("solid", fgColor=XL_GREEN_F),
    font=Font(name=FONT, size=10, bold=True, color=XL_GREEN_T)))
ws.conditional_formatting.add(f"D{DATA_START}:D{DATA_END}", CellIsRule(
    operator="equal", formula=['"被驳回"'], fill=PatternFill("solid", fgColor=XL_RED_F),
    font=Font(name=FONT, size=10, bold=True, color=XL_RED_T)))
ws.conditional_formatting.add(f"D{DATA_START}:D{DATA_END}", CellIsRule(
    operator="equal", formula=['"审核中"'], fill=PatternFill("solid", fgColor=XL_YEL_F),
    font=Font(name=FONT, size=10, bold=True, color=XL_YEL_T)))

ws.freeze_panes = "D4"
ws.auto_filter.ref = f"A3:J{DATA_END}"

def autofit_rows(ws, start_row=1, max_h=160):
    """按列宽估算换行后的行数，给出不小于内容所需的行高（避免文字被裁切）。"""
    merged = list(ws.merged_cells.ranges)
    for r in range(start_row, ws.max_row + 1):
        lines = 1
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if not isinstance(cell.value, str) or not cell.value.strip():
                continue
            w = ws.column_dimensions[get_column_letter(c)].width or 8.43
            for mr in merged:
                if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                    w = sum((ws.column_dimensions[get_column_letter(k)].width or 8.43)
                            for k in range(mr.min_col, mr.max_col + 1))
                    break
            cpl = max(4, int(w / 2.0))
            text_lines = sum(max(1, -(-len(seg) // cpl)) for seg in str(cell.value).split("\n"))
            lines = max(lines, text_lines)
        h = min(max_h, max(18, lines * 13 + 6))
        cur = ws.row_dimensions[r].height
        if cur is None or h > cur:
            ws.row_dimensions[r].height = h


for _ws in wb.worksheets:
    autofit_rows(_ws, start_row=3 if _ws.title != "08_发布进度" else 4)

# 打印：一律横向 + 按宽度缩放到 1 页，避免列被切到第二页
for _ws in wb.worksheets:
    _ws.page_setup.orientation = "landscape"
    _ws.page_setup.fitToWidth = 1
    _ws.page_setup.fitToHeight = 0
    _ws.sheet_properties.pageSetUpPr.fitToPage = True
    _ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    _ws.print_options.horizontalCentered = True

wb.save(OUT)
print("saved:", OUT)
print("sheets:", wb.sheetnames)
