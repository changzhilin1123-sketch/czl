# -*- coding: utf-8 -*-
"""Generate live streaming equipment rental comparison Excel (no price version)."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT = Path(__file__).resolve().parent / "直播租赁六家对比表_无价格版.xlsx"

HEADER_FILL = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name="微软雅黑", bold=True, size=14, color="2E7D32")
BODY_FONT = Font(name="微软雅黑", size=10)
NOTE_FONT = Font(name="微软雅黑", size=9, color="666666")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
ALT_FILL = PatternFill(start_color="F1F8F2", end_color="F1F8F2", fill_type="solid")


def style_header_row(ws, row, col_count):
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def style_data_area(ws, start_row, end_row, col_count, alt=False):
    for r in range(start_row, end_row + 1):
        for c in range(1, col_count + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if c == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if alt and (r - start_row) % 2 == 1:
                cell.fill = ALT_FILL


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_title(ws, title, note=None, merge_cols=10):
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=merge_cols)
    row = 2
    if note:
        ws[f"A{row}"] = note
        ws[f"A{row}"].font = NOTE_FONT
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_cols)
        row += 1
    return row


def write_table(ws, start_row, headers, rows, col_widths):
    for c, h in enumerate(headers, 1):
        ws.cell(row=start_row, column=c, value=h)
    style_header_row(ws, start_row, len(headers))
    for i, row_data in enumerate(rows):
        r = start_row + 1 + i
        for c, val in enumerate(row_data, 1):
            ws.cell(row=r, column=c, value=val)
    end_row = start_row + len(rows)
    style_data_area(ws, start_row + 1, end_row, len(headers), alt=True)
    set_col_widths(ws, col_widths)
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    return end_row + 2


def build_workbook():
    wb = Workbook()

    # ===== Sheet 1: 总览 =====
    ws1 = wb.active
    ws1.title = "总览"
    r = write_title(
        ws1,
        "直播租赁六家对比 — 配置总览（无价格版）",
        "租期：3个月起 | 服务：六家均含上门调试 + 日常远程协助 | 场景：武隆天生三桥户外 + 室内绿幕",
        merge_cols=11,
    )
    headers1 = [
        "商家", "机身", "镜头", "采集卡", "麦克风", "散热器",
        "是否含电脑", "电脑配置摘要", "户外网络", "户外电源", "绿幕", "其他亮点",
    ]
    rows1 = [
        ["租小杨", "Sony FX30", "索尼18-105", "数魅VC20", "猛玛一拖二", "无", "是",
         "i5-12400F/32G/512G/RTX4060/27寸1K", "无", "Power1000×1", "无", "64G卡+读卡器+备用电池；售后条款强"],
        ["昊天", "Sony A7C", "腾龙28-75", "数魅VC20", "猛玛一拖二", "无", "是",
         "i5-12400F/32G/512G/RTX3060 12G/27寸1K", "无", "无", "无", "顶光+魔术腿；3个月起租包邮寄出"],
        ["阿水", "Sony FX30", "适马18-50 F2.8", "圆刚553", "猛玛M1+充电盒", "有", "是",
         "i7 11代/32G/512G/RTX3060 6G 台式", "多网聚合路由", "Power1000×2", "无", "KK360W×4灯光；户外配件最全"],
        ["猫冬", "Sony A7M4", "腾龙28-75/适马28-70", "数魅VC20 Pro", "一拖二(品牌未写)", "有", "是",
         "i5-12400F/32G/512G/RTX3060 6G 台式", "无", "无", "无", "4×数魅300B；供电配件全"],
        ["洋洋", "Sony A7M4", "腾龙28-75/适马28-70", "数魅VC20 Pro", "一拖二(品牌未写)", "有", "是",
         "i5 12代/32G/512G/RTX3060/27寸 台式", "无", "无", "无", "印迹200D面光+300B；64G卡"],
        ["易希", "Sony A7M4", "适马28-70 F2.8", "圆刚553", "猛玛M1+充电盒", "有", "否(可另租)",
         "可参考同配置台式(i5/32G/3060/27寸)", "无", "无", "3×2.8m+龙门架(赠)", "唯一送绿幕；灯光分灯位适合抠像"],
    ]
    write_table(ws1, r, headers1, rows1, [10, 14, 18, 14, 16, 8, 10, 28, 12, 14, 8, 28])

    # ===== Sheet 2: 相机采集 =====
    ws2 = wb.create_sheet("相机与采集")
    r = write_title(ws2, "相机与采集配置对比")
    headers2 = [
        "商家", "机身", "镜头", "UV镜", "采集卡", "HDMI线",
        "假电池", "备用电池+充电器", "64G卡+读卡器", "散热器", "三脚架+竖拍板+包", "麦克风",
    ]
    rows2 = [
        ["租小杨", "Sony FX30", "索尼18-105", "无", "数魅VC20", "专用线", "有", "有", "有", "无", "有", "猛玛一拖二"],
        ["昊天", "Sony A7C", "腾龙28-75", "无", "数魅VC20", "2条", "有", "有", "有", "无", "有", "猛玛一拖二"],
        ["阿水", "Sony FX30", "适马18-50 F2.8", "无", "圆刚553", "5m×1", "有", "无", "无", "有", "有", "猛玛M1+充电盒"],
        ["猫冬", "Sony A7M4", "腾龙28-75/适马28-70", "有", "数魅VC20 Pro", "绿联5m", "有", "有", "无", "有", "有", "一拖二(品牌未写)"],
        ["洋洋", "Sony A7M4", "腾龙28-75/适马28-70", "有", "数魅VC20 Pro", "绿联5m", "有", "有", "有", "有", "有", "一拖二(品牌未写)"],
        ["易希", "Sony A7M4", "适马28-70 F2.8", "无", "圆刚553", "5m×1", "有", "无", "无", "有", "有", "猛玛M1+充电盒"],
    ]
    write_table(ws2, r, headers2, rows2, [10, 14, 18, 8, 14, 10, 8, 14, 14, 8, 16, 16])

    # ===== Sheet 3: 灯光配件 =====
    ws3 = wb.create_sheet("灯光与配件")
    r = write_title(ws3, "灯光与场景配件对比")
    headers3 = ["商家", "灯具", "单灯功率", "灯位/数量", "柔光罩", "灯架", "顶光", "绿幕", "户外电源", "户外网络"]
    rows3 = [
        ["租小杨", "捷宝/凛光/数魅200W", "200W", "数量未详列", "抛物线+球形等", "有", "未列", "无", "Power1000×1", "无"],
        ["昊天", "捷宝200W×4", "200W", "2单色+2双色", "90cm抛物线+球形+方形等", "4个", "有(顶光+魔术腿)", "无", "无", "无"],
        ["阿水", "KK 360WBi×4", "360W", "辅光/底子/造型/辅光", "90cm深抛物线+65cm灯笼+60×90方形", "4×2.8m", "无", "无", "Power1000×2", "多网聚合路由"],
        ["猫冬", "数魅300B×4", "300W", "2面光+2环境光", "90cm深口抛物线+65cm球形", "4×2.8m", "无", "无", "无", "无"],
        ["洋洋", "印迹200D×2+数魅300B×2", "200W/300W", "2面光+2环境光(4灯)", "90cm深口抛物线+65cm球形", "4×2.8m", "无", "无", "无", "无"],
        ["易希", "200W×4(登伟/数魅随机)", "200W", "主光/底子/特写/辅光", "90cm抛物线+65cm球形+60×90方形", "4×2.8m", "无", "3×2.8m+龙门架(赠)", "无", "无"],
    ]
    write_table(ws3, r, headers3, rows3, [10, 22, 10, 18, 24, 10, 14, 18, 14, 14])

    # ===== Sheet 4: 推流电脑 =====
    ws4 = wb.create_sheet("推流电脑")
    r = write_title(ws4, "推流电脑配置对比（不含租金）")
    headers4 = [
        "商家", "是否含电脑", "电脑类型", "CPU", "内存", "硬盘",
        "显卡", "显示器", "适用说明",
    ]
    rows4 = [
        ["租小杨", "是", "台式", "i5-12400F", "32G", "512G SSD", "RTX4060", "27寸1K", "4060推流+美颜余量最大"],
        ["昊天", "是", "台式", "i5-12400F", "32G", "512G SSD", "RTX3060 12G", "27寸1K", "12G显存适合多图层/高码率"],
        ["猫冬", "是", "台式", "i5-12400F", "32G", "512G SSD", "RTX3060 6G", "未标注", "直播推流标准配置"],
        ["洋洋", "是", "台式(单平台)", "i5 12代", "32G", "512G SSD", "RTX3060 6G", "27寸", "固定直播间推流"],
        ["阿水", "是", "台式", "i7 11代", "32G", "512G SSD", "RTX3060 6G", "未标注", "CPU略强+32G内存"],
        ["易希", "否(可另租)", "台式", "i5 12代(参考)", "32G", "512G SSD", "RTX3060 6G", "27寸(参考)", "需向商家另租或外配"],
    ]
    write_table(ws4, r, headers4, rows4, [10, 12, 14, 14, 8, 12, 14, 10, 24])

    # ===== Sheet 5: 服务支持 =====
    ws5 = wb.create_sheet("服务支持")
    r = write_title(ws5, "服务与支持对比", "上门调试与远程协助：六家均已沟通确认提供")
    headers5 = [
        "商家", "最短租期", "上门调试", "日常远程协助", "物流/包邮",
        "非人为免费维修", "押金/退押政策", "售后时段", "其他说明",
    ]
    rows5 = [
        ["租小杨", "3个月", "是", "是", "需确认", "有(非人为)", "满3月退押无违约", "8:30-22:30", "不限次调试+维保"],
        ["昊天", "3个月起租", "是", "是", "包邮寄出", "需确认", "需确认", "需确认", "含顶光+魔术腿"],
        ["阿水", "3个月", "是", "是", "单程包邮(报价单)", "需确认", "需确认", "需确认", "户外配件最全"],
        ["猫冬", "3个月", "是", "是", "需确认", "需确认", "需确认", "需确认", "A7M4+300B方案"],
        ["洋洋", "3个月", "是", "是", "需确认", "需确认", "需确认", "需确认", "A7M4+印迹面光方案"],
        ["易希", "3个月", "是", "是", "单程包邮(报价单)", "需确认", "需确认", "需确认", "送绿幕"],
    ]
    write_table(ws5, r, headers5, rows5, [10, 12, 10, 14, 12, 14, 16, 12, 22])

    # ===== Sheet 6: 场景匹配 =====
    ws6 = wb.create_sheet("场景匹配建议")
    r = write_title(
        ws6,
        "场景匹配 — 武隆户外 + 室内绿幕 + 高画质",
        "本表仅对比配置与场景适配，不含价格信息",
    )
    headers6 = [
        "商家", "画质档", "户外网络", "户外电源", "绿幕", "含电脑",
        "电脑显卡", "灯光特点", "主要优势", "主要缺口",
    ]
    rows6 = [
        ["租小杨", "高(FX30)", "无", "Power1000×1", "无", "是", "RTX4060",
         "200W混用", "FX30长播+4060+户外电源+强售后", "无绿幕/聚合路由；18-105光圈小"],
        ["昊天", "中(A7C)", "无", "无", "无", "是", "RTX3060 12G",
         "200W+顶光", "顶光高级感+12G显存+包邮寄出", "无绿幕/户外电源/路由；机身档略低"],
        ["阿水", "高(FX30+360W)", "有", "双Power1000", "无", "是", "RTX3060 6G",
         "360W×4", "户外网络+电源+灯光最全", "无绿幕；18-50远景弱"],
        ["猫冬", "高(A7M4+300B)", "无", "无", "无", "是", "RTX3060 6G",
         "300B×4", "A7M4+VC20Pro+均匀布光", "无绿幕/户外配件；麦品牌未写"],
        ["洋洋", "高(A7M4+印迹)", "无", "无", "无", "是", "RTX3060 6G",
         "印迹200D+300B", "肤色/面光显色好+64G卡", "无绿幕/户外配件；麦品牌未写"],
        ["易希", "高(A7M4)", "无", "无", "有(赠)", "否(可另租)", "需另配",
         "200W分灯位", "唯一送绿幕+适合抠像", "电脑需另租；无户外配件；灯品牌随机"],
    ]
    write_table(ws6, r, headers6, rows6, [10, 14, 10, 14, 12, 10, 12, 16, 24, 24])

    # ===== Sheet 7: 优缺点 =====
    ws7 = wb.create_sheet("各家优缺点")
    r = write_title(ws7, "各家优缺点（配置维度 · 无价格版）",
                    "针对：3个月 / 含电脑推流 / 武隆户外+室内绿幕 / 高画质")
    headers7 = ["商家", "套餐定位", "优点", "缺点", "综合适合度"]
    rows7 = [
        [
            "租小杨", "FX30拍摄+4060电脑+户外电源",
            "4060显卡推流余量最大；FX30适合长播；Power1000解决户外供电；配件全(64G卡/备用电池)；售后条款最强",
            "无绿幕/聚合路由；18-105光圈偏小弱光弱；无散热器；灯光数量未详",
            "★★★★★ 户外有电源+强显卡+强售后",
        ],
        [
            "昊天", "A7C拍摄+3060 12G电脑+顶光",
            "顶光+魔术腿提升直播间质感；3060 12G显存充裕；配件齐全；3个月起租包邮寄出",
            "A7C画质档低于A7M4/FX30；无绿幕/户外电源/路由；无散热",
            "★★★★ 固定直播间+重视顶光效果",
        ],
        [
            "阿水", "FX30拍摄+360W灯光+户外全套+电脑",
            "唯一含聚合路由+双Power1000；KK360W灯光功率最高；FX30+猛玛M1；32G台式",
            "无绿幕；18-50镜头远景/虚化弱于28-70；户外与室内需切换布光",
            "★★★★★ 户外直播为主、要求最省心",
        ],
        [
            "猫冬", "A7M4拍摄+300B灯光+3060电脑",
            "A7M4+VC20 Pro采集；4×300B布光均匀；供电配件全(假电池+备用+充电器)",
            "无绿幕/路由/电源；麦克风品牌未写；台式不便携",
            "★★★★ 室内高画质+标准全套",
        ],
        [
            "洋洋", "A7M4拍摄+印迹面光+3060电脑",
            "A7M4+VC20 Pro；印迹200D高显色面光；64G卡；i5 12代+3060+27寸显示器",
            "无绿幕/路由/电源；麦克风品牌未写；与猫冬同档但灯光方案不同",
            "★★★★ 重视肤色与面光质量",
        ],
        [
            "易希", "A7M4拍摄+绿幕+分灯位(电脑另租)",
            "唯一送绿幕；A7M4+553采集；主/辅/底子/特写四灯位适合绿幕抠像；有散热",
            "电脑需另租；200W灯且品牌随机；无户外网络/电源；仅假电池",
            "★★★★ 室内绿幕比重大",
        ],
    ]
    write_table(ws7, r, headers7, rows7, [10, 22, 38, 38, 24])

    # ===== Sheet 8: 配置差异速查 =====
    ws8 = wb.create_sheet("配置差异速查")
    r = write_title(ws8, "六家关键配置差异速查（无价格）")
    headers8 = ["对比维度", "租小杨", "昊天", "阿水", "猫冬", "洋洋", "易希"]
    rows8 = [
        ["机身档次", "FX30(影视)", "A7C(轻便)", "FX30(影视)", "A7M4(高)", "A7M4(高)", "A7M4(高)"],
        ["镜头特点", "18-105变焦", "28-75 F2.8", "18-50 F2.8广角", "28-75/28-70", "28-75/28-70", "28-70 F2.8"],
        ["采集卡", "VC20", "VC20", "圆刚553", "VC20 Pro", "VC20 Pro", "圆刚553"],
        ["灯光功率", "200W级", "200W+顶光", "360W(最高)", "300W", "200D+300B", "200W"],
        ["户外电源", "Power×1", "无", "Power×2", "无", "无", "无"],
        ["户外网络", "无", "无", "聚合路由", "无", "无", "无"],
        ["绿幕", "无", "无", "无", "无", "无", "赠送"],
        ["电脑显卡", "RTX4060", "RTX3060 12G", "RTX3060 6G", "RTX3060 6G", "RTX3060 6G", "需另配"],
        ["电脑内存", "32G", "32G", "32G", "32G", "32G", "需另配"],
        ["显示器", "27寸1K", "27寸1K", "未标注", "未标注", "27寸", "需另配"],
        ["包邮政策", "需确认", "包邮寄出", "单程包邮", "需确认", "需确认", "单程包邮"],
    ]
    write_table(ws8, r, headers8, rows8, [14, 16, 16, 16, 16, 16, 16])

    wb.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    path = build_workbook()
    print(f"已生成: {path}")
