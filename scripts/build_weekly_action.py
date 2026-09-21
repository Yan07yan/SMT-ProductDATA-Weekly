# -*- coding: utf-8 -*-
"""《周度行动表》生成器 —— 对标公司 KPI，输出本周该干什么。

输出：速卖通_周度行动表_YYYY-MM-DD.xlsx（3 页）
  页1 公司 KPI 与本周缺口（销售额/刊登数/热销链接/全托管订单/滞销清理/利润率）
  页2 本周行动清单（KPI 缺口拆解 + 产品优化 TOP10）
  页3 数据底稿（流量/转化/滞销明细，支撑页1-2）
  页4 上周好品类与本周方向（上周卖得好的品类 + 本周上品方向草稿/手填）

用法：
  python build_weekly_action.py --data <每周导出目录> [--date 2026-09-14]
                                 [--week-start 2026-09-08 --week-end 2026-09-12]
                                 [--out 输出.xlsx] [--track 追踪表.xlsx] [--snapshot 快照.json]

配置（账号专属，放 config.json，不要改代码）：
  - 复制 config.example.json 为 config.json，填入你自己的：
      · saas_shop：订单文件里的渠道名 → 店铺代号映射
      · shop_file / cat_file / cat_week_file：店铺代号 → 导出文件名
      · kpi：本公司当月月度考核目标
  - 脚本按优先级读配置：--config > <skill根>/config.json > config.example.json > 内置占位。
  - 列名取自速卖通后台标准导出，跨账号通用；文件名遵循「店号-30天数据.xlsx」等约定。
  - 双轨：30天产品文件(*-30天数据.xlsx)算 KPI；另需导出「上周窗口」产品文件(*-上周数据.xlsx)
    专供页4「上周好品类/本周上品方向」。缺上周文件时自动回退30天文件并标⚠️。
"""
import sys, os, json, argparse, statistics, datetime
from collections import defaultdict
import openpyxl, xlrd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============ 模块级配置（由 main 根据命令行参数填充） ============
DATA = None                       # 每周导出数据目录
OUT = None                        # 输出 xlsx
SNAP = None                       # 上品ID快照 json（跨周持久化，勿删）
LISTING_TRACK_FILE = None        # 刊登数量追踪表 xlsx
TODAY = None                      # 报告日期
WEEK_START = None                 # 刊登追踪汇总窗口起
WEEK_END = None                   # 刊登追踪汇总窗口止
CAT_WEEK_USED_FALLBACK = False       # 上周品类文件缺失时回退30天文件（报告中标⚠️）

PLAN_PER_WEEK = 50       # POP 周计划上架总数（大小周 6 天改 60）
TOPN = 10                # 每周优先改的数量

# 内置占位配置：不含任何真实店铺 ID / 订单渠道名 / 考核数字。
# 实际值从 config.json 读取（见 load_config）。同事接入只改配置、不动本文件。
DEFAULT_CONFIG = {
    "saas_shop": {
        "Aliexpress-<你的店铺1渠道名>": "SHOP_1",
        "Aliexpress-<你的店铺2渠道名>": "SHOP_2",
        "Aliexpress-<你的全托管渠道名>": "CLYT1",
    },
    "shop_file": {
        "SHOP_1": "SHOP1-30天数据.xlsx",
        "SHOP_2": "SHOP2-30天数据.xlsx",
    },
    "cat_file": {
        "SHOP_1": "SHOP1商品总数据.xlsx",
        "SHOP_2": "SHOP2商品总数据.xlsx",
    },
    "cat_week_file": {
        "SHOP_1": "SHOP1-上周数据.xlsx",
        "SHOP_2": "SHOP2-上周数据.xlsx",
    },
    "kpi": {
        "销售额":    {"target": 100000, "unit": "CNY", "weight": 0.35, "note": "按本公司月度考核填"},
        "利润率":    {"target": 0.15,  "unit": "%",   "weight": 0.15, "note": "5分12%/10分15%"},
        "刊登数量":  {"target": 690,   "unit": "个",  "weight": 0.10, "note": "按本月工作日填"},
        "热销链接":  {"target": 18,    "unit": "个",  "weight": 0.10, "note": "出单≥15单"},
        "全托管订单": {"target": 1077, "unit": "单",  "weight": 0.15, "note": "按考核填"},
        "滞销清理":  {"target": 0.60,  "unit": "%",   "weight": 0.10, "note": "清理占滞销品比例"},
    },
}

# 模块级占位（main 加载配置后会被真实值覆盖）
SALES_SHOP = dict(DEFAULT_CONFIG["saas_shop"])
SHOP_FILE = {}
CAT_FILE = {}
CAT_WEEK_FILE = {}
KPI = dict(DEFAULT_CONFIG["kpi"])

# ============ 样式 ============
H_FILL = PatternFill("solid", fgColor="1F4E78"); H_FONT = Font(color="FFFFFF", bold=True, size=11)
B_FILL = PatternFill("solid", fgColor="DDEBF7"); B_FONT = Font(bold=True)
T_FILL = PatternFill("solid", fgColor="FCE4D6"); Y_FILL = PatternFill("solid", fgColor="FFF2CC")
G_FILL = PatternFill("solid", fgColor="E2EFDA"); R_FILL = PatternFill("solid", fgColor="F8CBAD")
SUB_F = Font(italic=True, size=9, color="666666")
TITLE_F = Font(bold=True, size=14, color="1F4E78")
WRAP = Alignment(wrap_text=True, vertical="top"); CEN = Alignment(horizontal="center", vertical="center")
thin = Side(style="thin", color="BFBFBF"); BORD = Border(left=thin, right=thin, top=thin, bottom=thin)


def put(ws, r, vals, fill=None, font=None, fmts=None, aligns=None):
    for i, v in enumerate(vals):
        c = ws.cell(row=r, column=i + 1, value=v); c.border = BORD
        if fill: c.fill = fill
        if font: c.font = font
        if fmts and i < len(fmts) and fmts[i]: c.number_format = fmts[i]
        if aligns and i < len(aligns) and aligns[i]: c.alignment = aligns[i]
        else: c.alignment = WRAP
    return r + 1


def setw(ws, widths):
    for i, w in enumerate(widths): ws.column_dimensions[get_column_letter(i + 1)].width = w


# ==================== 解析 ====================
def colidx(header, *names):
    for n in names:
        if n in header: return header.index(n)
    for n in names:
        for i, h in enumerate(header):
            if h and n in h: return i
    return None


def parse_30():
    out = {}
    for shop, f in SHOP_FILE.items():
        p = os.path.join(DATA, f)
        if not os.path.exists(p): continue
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        if "汇总指标" not in wb.sheetnames:
            wb.close(); continue
        ws = wb["汇总指标"]; rows = list(ws.iter_rows(values_only=True))
        hdr = [("" if c is None else str(c).strip()) for c in rows[0]]
        ci = lambda *n: colidx(hdr, *n)
        idx = {k: ci(v) for k, v in {"pid": "商品ID", "expose": "搜索曝光量", "ctr": "搜索点击率",
            "uv": "访客数", "cvr": "支付转化率", "views": "浏览量", "amt": "成交金额",
            "cat": "根类目名称", "orders": "支付商品件数", "fav": "商品收藏人数", "cart": "商品加购人数"}.items()}
        items = []
        for r in rows[1:]:
            if not r or r[0] is None: continue
            pid = str(r[idx["pid"]]).strip()
            if not (pid.isdigit() and len(pid) >= 10): continue
            def g(k):
                i = idx[k]; return r[i] if i is not None and i < len(r) else None
            def f(x):
                try: return float(x)
                except: return 0.0
            items.append({"pid": pid, "shop": shop, "expose": f(g("expose")), "ctr": f(g("ctr")),
                "uv": f(g("uv")), "cvr": f(g("cvr")), "views": f(g("views")), "amt": f(g("amt")),
                "cat": (str(g("cat")) if g("cat") else ""), "orders": f(g("orders")),
                "fav": f(g("fav")), "cart": f(g("cart"))})
        out[shop] = items; wb.close()
    return out


def parse_ft():
    """解析 CLYT1 全托 30 天数据（汇总指标 sheet）。"""
    if not os.path.exists(FT_FILE): return []
    wb = openpyxl.load_workbook(FT_FILE, read_only=True, data_only=True)
    if "汇总指标" not in wb.sheetnames:
        wb.close(); return []
    ws = wb["汇总指标"]; rows = list(ws.iter_rows(values_only=True))
    hdr = [("" if c is None else str(c).strip()) for c in rows[0]]
    ci = lambda *n: colidx(hdr, *n)
    idx = {"pid": ci("商品ID"), "orders": ci("支付订单数"), "pieces": ci("支付件数"),
         "amt": ci("支付金额", "成交金额"),
         "expose": ci("搜索曝光量"), "ctr": ci("搜索点击率"), "uv": ci("商品访客数"), "cvr": ci("支付转化率")}
    items = []
    for r in rows[1:]:
        if not r or r[0] is None: continue
        def g(k): i = idx[k]; return r[i] if i is not None and i < len(r) else None
        def f(x):
            try: return float(x)
            except: return 0.0
        pid = str(g("pid") or "").strip()
        if not (pid.isdigit() and len(pid) >= 10): continue
        orders = f(g("orders")) if idx["orders"] is not None else f(g("pieces"))
        items.append({"pid": pid, "orders": orders, "amt": f(g("amt")),
                      "expose": f(g("expose")), "ctr": f(g("ctr")), "uv": f(g("uv")), "cvr": f(g("cvr"))})
    wb.close(); return items


def parse_sales():
    """五店「销售额订单数据」.xls —— 订单级明细。
    返回 (res, win)：res={渠道:{orders,amt,cost,qty}}；win=(最早日期,最晚日期) 或 None。
    注意：全托管(QTG)行 产品总金额=0，销售额/毛利仅对四店 POP 有效。"""
    if not os.path.exists(SALES_FILE): return {}, None
    b = xlrd.open_workbook(SALES_FILE)
    s = b.sheet_by_index(0)
    hdr = [str(s.cell_value(0, c)).strip() for c in range(s.ncols)]
    ci = {n: hdr.index(n) for n in hdr}
    def num(x):
        try:
            if x is None or x == "": return 0.0
            return float(x)
        except: return 0.0
    chan_i = ci["订单来源渠道"]; amt_i = ci["产品总金额"]; qty_i = ci["发货产品数量"]
    sup_i = ci["发货产品供货价"]; add_i = ci["添加时间"]
    res = defaultdict(lambda: {"orders": 0, "amt": 0.0, "cost": 0.0, "qty": 0.0})
    dates = []
    for r in range(1, s.nrows):
        chan = str(s.cell_value(r, chan_i)).strip()
        amt = num(s.cell_value(r, amt_i)); qty = num(s.cell_value(r, qty_i)); sup = num(s.cell_value(r, sup_i))
        res[chan]["orders"] += 1; res[chan]["amt"] += amt; res[chan]["qty"] += qty
        res[chan]["cost"] += sup * (qty if qty else 1)
        dates.append(str(s.cell_value(r, add_i)).strip())
    ds = []
    for d in dates:
        for fmt in ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
            try: ds.append(datetime.datetime.strptime(d, fmt)); break
            except: pass
    win = (min(ds).date(), max(ds).date()) if ds else None
    return dict(res), win


def parse_ft_listing():
    """全托管上架数据.xlsx —— 按创建时间过滤上周上架。返回数量。文件若为空返回 0。"""
    if not os.path.exists(FT_LIST_FILE): return 0
    try:
        wb = openpyxl.load_workbook(FT_LIST_FILE, read_only=True, data_only=True)
    except: return 0
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        wb.close(); return 0
    hdr = [("" if c is None else str(c).strip()) for c in rows[0]]
    ci = lambda n: hdr.index(n) if n in hdr else None
    tcol = ci("创建时间") or ci("创建日期") or ci("上架时间") or ci("添加时间") or ci("上传时间")
    n = 0
    for r in rows[1:]:
        if tcol is None or tcol >= len(r): continue
        t = str(r[tcol] or "").strip()
        dt = None
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
            try: dt = datetime.datetime.strptime(t[:19], fmt); break
            except: pass
        if dt and WEEK_START <= dt.date() <= WEEK_END:
            n += 1
    wb.close(); return n


def parse_settle_sales():
    """全托结算文件 → (实际结算金额合计, 销售件数合计)。全托销售额近似=实际结算金额(净)。"""
    if not os.path.exists(SETTLE_FILE): return 0.0, 0
    wb = openpyxl.load_workbook(SETTLE_FILE, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [("" if c is None else str(c).strip()) for c in rows[0]]
    ci = lambda n: hdr.index(n) if n in hdr else None
    amt_i = ci("实际结算金额"); qty_i = ci("商品销售数量")
    def num(x):
        try: return float(x) if x not in (None, "") else 0.0
        except: return 0.0
    amt = 0.0; qty = 0
    for r in rows[1:]:
        amt += num(r[amt_i]) if amt_i is not None else 0
        qty += num(r[qty_i]) if qty_i is not None else 0
    wb.close()
    return amt, qty


def parse_listing_tracker(path=None, week_start=None, week_end=None):
    """刊登数量追踪.xlsx —— 用户每日维护的上品数台账。
    返回 {"total":x,"pop":y,"ft":z}；缺文件返回全 0。
    week_start/week_end 为 datetime.date；未传则使用模块全局 WEEK_START/WEEK_END。
    CLYT1 列 = 全托管；YZ3/WHDS06/PHCX04/ZHHK04 = POP 四店。"""
    path = path or LISTING_TRACK_FILE
    if not os.path.exists(path): return {"total": 0, "pop": 0, "ft": 0}
    if week_start is None: week_start = WEEK_START
    if week_end is None: week_end = WEEK_END
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2: wb.close(); return {"total": 0, "pop": 0, "ft": 0}
    hdr = [("" if c is None else str(c).strip()) for c in rows[0]]
    ci = lambda n: hdr.index(n) if n in hdr else None
    date_i = colidx(hdr, "日期", "Date", "date", "时间")
    cols = {"CLYT1": ci("CLYT1"), "YZ3": ci("YZ3"), "WHDS06": ci("WHDS06"),
            "PHCX04": ci("PHCX04"), "ZHHK04": ci("ZHHK04")}
    pop_shops = ("YZ3", "WHDS06", "PHCX04", "ZHHK04")
    total = pop = ft = 0
    for r in rows[1:]:
        if date_i is None or date_i >= len(r): continue
        dv = r[date_i]
        dt = None
        if isinstance(dv, datetime.datetime): dt = dv.date()
        elif isinstance(dv, datetime.date): dt = dv
        elif isinstance(dv, (int, float)) and dv > 30000:
            try: dt = (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(dv))).date()
            except: pass
        else:
            t = str(dv or "").strip()
            parts = [p.strip() for p in t.replace("-", "/").split("/") if p.strip()]
            if len(parts) == 3:
                try:
                    y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                    dt = datetime.date(y, m, d)
                except: pass
            if dt is None:
                for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
                    try: dt = datetime.datetime.strptime(t[:19], fmt).date(); break
                    except: pass
        if dt is None or not (week_start <= dt <= week_end): continue
        def num(v):
            try: return float(v) if v not in (None, "", "/") else 0.0
            except: return 0.0
        ft += num(r[cols["CLYT1"]]) if cols["CLYT1"] is not None else 0
        for s in pop_shops:
            i = cols[s]
            pop += num(r[i]) if i is not None else 0
    total = pop + ft
    wb.close()
    return {"total": total, "pop": pop, "ft": ft}


def parse_catalog():
    out = {}
    for shop, f in CAT_FILE.items():
        p = os.path.join(DATA, f)
        if not os.path.exists(p): continue
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            hi = None
            for ri, r in enumerate(rows[:5]):
                if any(r and str(c).strip() == "id" for c in r if c is not None): hi = ri; break
            if hi is None: continue
            hdr = [("" if c is None else str(c).strip()) for c in rows[hi]]
            ci = lambda *n: colidx(hdr, *n)
            ti = ci("商品标题", "title"); idi = ci("id")
            wi = ci("物流重量", "logisticsWeight"); si = ci("日销运费模版", "freightTemplate")
            ei = ci("关联欧盟责任人", "msrEuComp")
            for r in rows[hi + 1:]:
                if not r or r[idi] is None: continue
                pid = str(r[idi]).strip()
                if not (pid.isdigit() and len(pid) >= 10): continue
                def g(i): return r[i] if i is not None and i < len(r) else None
                out[pid] = {"title": (str(g(ti))[:40] if g(ti) else ""),
                            "weight": (float(g(wi)) if g(wi) else 0),
                            "ship_tpl": (str(g(si))[:20] if g(si) else ""),
                            "eu": (str(g(ei))[:20] if g(ei) else "")}
        wb.close()
    return out


def parse_listing():
    if not os.path.exists(LIST_FILE): return {}, set()
    wb = xlrd.open_workbook(LIST_FILE); ws = wb.sheet_by_index(0)
    hdr = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]
    ci = lambda n: hdr.index(n) if n in hdr else None
    cols = {"chan": ci("来源渠道"), "item": ci("ItemId"), "cn": ci("产品中文名称"),
            "en": ci("产品英文名称"), "price": ci("售价"), "rmb": ci("人民币售价"),
            "d30": ci("30天销量"), "d90": ci("90天销量"), "op": ci("上架人员"), "link": ci("产品销售链接")}
    by_shop = {}; all_ids = set()
    chan2shop = {"PHC": "PHCX04", "YZ3": "YZ3", "WHDS": "WHDS06", "ZHHK": "ZHHK04"}
    for r in range(1, ws.nrows):
        row = [ws.cell_value(r, c) for c in range(ws.ncols)]
        chan = str(row[cols["chan"]]) if cols["chan"] is not None else ""
        item = str(row[cols["item"]]).strip() if cols["item"] is not None else ""
        if not item or not item.isdigit(): continue
        all_ids.add(item)
        shop = next((v for k, v in chan2shop.items() if k in chan), None)
        if not shop: continue
        by_shop.setdefault(shop, []).append({"item": item,
            "cn": str(row[cols["cn"]])[:30] if cols["cn"] is not None else "",
            "price": row[cols["price"]] if cols["price"] is not None else 0,
            "rmb": row[cols["rmb"]] if cols["rmb"] is not None else 0,
            "d30": row[cols["d30"]] if cols["d30"] is not None else 0,
            "d90": row[cols["d90"]] if cols["d90"] is not None else 0,
            "op": str(row[cols["op"]]) if cols["op"] is not None else "",
            "link": str(row[cols["link"]]) if cols["link"] is not None else ""})
    return by_shop, all_ids


# ==================== 上周品类 ====================
def parse_cat_week():
    """上周窗口产品文件(*-上周数据.xlsx) → 按根类目名称聚合品类表现。
    缺失上周文件时回退到 *-30天数据.xlsx（置 CAT_WEEK_USED_FALLBACK 标⚠️）。
    仅用于『上周好品类 / 本周上品方向』，不参与 KPI 口径。"""
    global CAT_WEEK_USED_FALLBACK
    CAT_WEEK_USED_FALLBACK = False
    out = {}
    for shop in SHOP_FILE:
        wk = os.path.join(DATA, CAT_WEEK_FILE[shop])
        fb = os.path.join(DATA, SHOP_FILE[shop])
        p = wk if os.path.exists(wk) else (fb if os.path.exists(fb) else None)
        if p is None:
            out[shop] = {}; continue
        if p == fb:
            CAT_WEEK_USED_FALLBACK = True
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        if "汇总指标" not in wb.sheetnames:
            wb.close(); out[shop] = {}; continue
        ws = wb["汇总指标"]; rows = list(ws.iter_rows(values_only=True))
        hdr = [("" if c is None else str(c).strip()) for c in rows[0]]
        ci = lambda *n: colidx(hdr, *n)
        idx = {"pid": ci("商品ID"), "cat": ci("根类目名称"),
               "orders": ci("支付商品件数"), "amt": ci("成交金额")}
        cats = defaultdict(lambda: {"orders": 0.0, "amt": 0.0, "n": 0})
        for r in rows[1:]:
            if not r or r[0] is None: continue
            pid = str(r[idx["pid"]]).strip()
            if not (pid.isdigit() and len(pid) >= 10): continue
            def num(x):
                try: return float(x) if x not in (None, "") else 0.0
                except: return 0.0
            cat = str(r[idx["cat"]]) if (idx["cat"] is not None and r[idx["cat"]] not in (None, "")) else "（未分类）"
            cats[cat]["orders"] += num(r[idx["orders"]]) if idx["orders"] is not None else 0.0
            cats[cat]["amt"] += num(r[idx["amt"]]) if idx["amt"] is not None else 0.0
            cats[cat]["n"] += 1
        wb.close(); out[shop] = dict(cats)
    return out


# ==================== KPI 计算 ====================
def med(vals):
    vals = [v for v in vals if v > 0]
    return statistics.median(vals) if vals else 0


def compute_kpis(d30, ft_items, by_shop, catalog, sales_res, sales_win, listing_counts):
    POP_SHOPS = ("ZHHK04", "YZ3", "WHDS06", "PHCX04")
    pop_sales = sum(v["amt"] for k, v in sales_res.items() if SALES_SHOP.get(k) in POP_SHOPS)
    ft_amt, _ = parse_settle_sales()
    win_days = (sales_win[1] - sales_win[0]).days + 1 if sales_win else 7
    monthly_runrate = pop_sales * (30.0 / win_days) + ft_amt if win_days > 0 else pop_sales + ft_amt

    pop_amt = 0.0; pop_cost = 0.0; store_m = {}
    for k, v in sales_res.items():
        shop = SALES_SHOP.get(k)
        if shop in POP_SHOPS and v["amt"] > 0:
            pop_amt += v["amt"]; pop_cost += v["cost"]
            store_m[shop] = {"gross": (v["amt"] - v["cost"]) / v["amt"],
                             "net": (v["amt"] - v["cost"]) / v["amt"] - 0.03,
                             "amt": v["amt"], "cost": v["cost"]}
    gross_m = (pop_amt - pop_cost) / pop_amt if pop_amt else 0
    net_m = gross_m - 0.03

    pop_listing = int(listing_counts.get("pop", 0))
    ft_list_count = int(listing_counts.get("ft", 0))
    listing_total = int(listing_counts.get("total", 0))

    hot_pop = sum(1 for shop in d30.values() for x in shop if x["orders"] >= 15)
    hot_ft = sum(1 for x in ft_items if x["orders"] >= 15)
    hotlinks = hot_pop + hot_ft

    ft_orders = sum(x["orders"] for x in ft_items)

    active_ids = {x["pid"] for shop in d30.values() for x in shop if x["orders"] > 0}
    all_ids = set(catalog.keys())
    if all_ids:
        dead_ids = all_ids - active_ids
        dead_count = len(dead_ids)
        clear_target = int(dead_count * KPI["滞销清理"]["target"])
    else:
        dead_ids = set(); dead_count = None; clear_target = None

    return {
        "销售额":   {"current": monthly_runrate, "pop": pop_sales, "ft": ft_amt, "win_days": win_days,
                     "target": KPI["销售额"]["target"]},
        "利润率":   {"current": net_m, "gross": gross_m, "store": store_m, "target": KPI["利润率"]["target"]},
        "刊登数量": {"current": listing_total, "pop": pop_listing, "ft": ft_list_count,
                     "target": KPI["刊登数量"]["target"]},
        "热销链接": {"current": hotlinks, "target": KPI["热销链接"]["target"]},
        "全托管订单": {"current": ft_orders, "target": KPI["全托管订单"]["target"]},
        "滞销清理": {"current": None, "target": clear_target, "denom": dead_count},
    }, dead_ids


# ==================== 行动清单 ====================
def build_actions(d30, catalog):
    main_img, detail, hotlink_push = [], [], []
    for shop, items in d30.items():
        m_ctr = med([x["ctr"] for x in items if x["expose"] >= 50])
        m_cvr = med([x["cvr"] for x in items if x["uv"] >= 50])
        for x in items:
            title = catalog.get(x["pid"], {}).get("title", "")
            rec = {"shop": shop, "pid": x["pid"], "title": title, "expose": x["expose"],
                   "ctr": x["ctr"], "uv": x["uv"], "cvr": x["cvr"], "orders": x["orders"],
                   "m_ctr": m_ctr, "m_cvr": m_cvr}
            if x["expose"] >= 300 and x["ctr"] < 0.7 * m_ctr:
                rec["lost"] = x["expose"] * max(0, m_ctr - x["ctr"])
                rec["diag"] = f"曝光{x['expose']:.0f}·点击率{x['ctr']*100:.1f}%<店均{m_ctr*100:.1f}%"
                rec["act"] = "重拍/换主图：加使用场景+等比尺寸参照；同步优化搜索卡片标题前20字符"
                main_img.append(rec)
            if x["uv"] >= 100 and x["cvr"] < 0.7 * m_cvr:
                rec2 = rec.copy(); rec2["lost_o"] = x["uv"] * max(0, m_cvr - x["cvr"])
                rec2["diag"] = f"访客{x['uv']:.0f}·转化率{x['cvr']*100:.1f}%<店均{m_cvr*100:.1f}%"
                rec2["act"] = "补详情页卖点/尺寸表/安装图，查价格差距，加视频/问答；先优化标题词"
                detail.append(rec2)
            if 8 <= x["orders"] < 15 and x["cvr"] >= 0.6 * m_cvr:
                rec3 = rec.copy(); rec3["gap_o"] = 15 - x["orders"]
                rec3["diag"] = f"当前出单{x['orders']:.0f}件，距热销线(15单)差{rec3['gap_o']:.0f}单"
                rec3["act"] = "冲热销链接：报平台活动/限时限量/满件优惠/私域定向，7天内推到15单"
                hotlink_push.append(rec3)
    main_img.sort(key=lambda r: r.get("lost", 0), reverse=True)
    detail.sort(key=lambda r: r.get("lost_o", 0), reverse=True)
    hotlink_push.sort(key=lambda r: r.get("gap_o", 0))
    return main_img, detail, hotlink_push


def listing_rate(by_shop, all_ids):
    cur = {s: set(x["item"] for x in v) for s, v in by_shop.items()}
    cur_all = all_ids
    prior = {}
    if os.path.exists(SNAP):
        try: prior = json.load(open(SNAP, encoding="utf-8"))
        except: prior = {}
    prev_week = prior.get("week")
    has_prev = bool(prev_week) and (prev_week != TODAY.isoformat())
    baseline_week = not has_prev
    result = {}
    for shop in SHOP_FILE:
        c = len(cur.get(shop, set())); p = len(prior.get("shops", {}).get(shop, []))
        new = max(0, c - p); rate = (new / PLAN_PER_WEEK) if has_prev else None
        result[shop] = {"online": c, "prior": p, "new": new, "rate": rate}
    total_online = len(cur_all); total_prior = len(set(prior.get("all", [])))
    total_new = max(0, total_online - total_prior)
    total_rate = (total_new / PLAN_PER_WEEK) if has_prev else None
    json.dump({"week": TODAY.isoformat(), "shops": {s: list(v) for s, v in cur.items()},
               "all": list(cur_all)}, open(SNAP, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return result, total_online, total_prior, total_new, total_rate, baseline_week


# ==================== 写表 ====================
# ==================== 上周好品类 / 本周上品方向 ====================
def recommend_cat_action(cat, v, rank, pct):
    if rank == 0:
        return f"头部品类(占{pct:.0f}%)，本周优先上架该品类新款/变体，巩固基本盘"
    if v["orders"] >= 20:
        return "稳健品类：保持上架节奏，可加套装提客单"
    if v["orders"] >= 5:
        return "成长品类：试上 2-3 个新款测转化"
    return "长尾品类：暂不强推，有库存可顺带铺"


def build_cat_week_sheet(wb, cat_week):
    """页4：上周卖得好的品类（按根类目名称聚合）+ 本周上品方向（草稿+手填列）。"""
    ws = wb.create_sheet("4_上周好品类与本周方向")
    setw(ws, [10, 24, 12, 14, 12, 46])
    ws.merge_cells("A1:F1")
    ws["A1"] = "上周卖得好的品类 & 本周上品方向"
    ws["A1"].font = TITLE_F
    ws.merge_cells("A2:F2")
    if CAT_WEEK_USED_FALLBACK:
        src = "⚠️ 数据源回退到 *-30天数据.xlsx（未导出上周文件，数字为30天口径非真实上周）"
    else:
        src = "数据源：*-上周数据.xlsx（上周窗口导出）"
    ws["A2"] = f"生成 {TODAY} ｜ {src} ｜ 品类=根类目名称；本周方向=数据草稿+你手填"
    ws["A2"].font = SUB_F
    r = 4
    any_data = False
    for shop in SHOP_FILE:
        cats = cat_week.get(shop, {})
        if not cats:
            ws.merge_cells(f"A{r}:F{r}")
            ws[f"A{r}"] = f"【{shop}】⚠️ 无数据（上周/30天产品文件缺失）"
            ws[f"A{r}"].fill = R_FILL; r += 2; continue
        any_data = True
        total_o = sum(v["orders"] for v in cats.values())
        ranked = sorted(cats.items(), key=lambda kv: kv[1]["orders"], reverse=True)
        ws.merge_cells(f"A{r}:F{r}")
        ws[f"A{r}"] = f"【{shop}】上周好品类 TOP{min(6, len(ranked))}（共 {len(cats)} 个品类，总支付 {total_o:.0f} 件）"
        ws[f"A{r}"].font = B_FONT; ws[f"A{r}"].fill = Y_FILL; r += 1
        r = put(ws, r, ["排名", "品类", "支付件数", "成交金额", "占店比", "数据建议"],
                fill=H_FILL, font=H_FONT, aligns=[CEN, None, CEN, CEN, CEN, None])
        for i, (cat, v) in enumerate(ranked[:6]):
            pct = (v["orders"] / total_o * 100) if total_o else 0
            r = put(ws, r, [i + 1, cat, f"{v['orders']:.0f}", f"¥{v['amt']:,.0f}",
                            f"{pct:.0f}%", recommend_cat_action(cat, v, i, pct)],
                    aligns=[CEN, None, CEN, CEN, CEN, None])
        r += 1
        ws.merge_cells(f"A{r}:F{r}")
        ws[f"A{r}"] = f"▸ {shop} 本周上品方向（草稿，最后一列★你手填）"
        ws[f"A{r}"].font = B_FONT; ws[f"A{r}"].fill = T_FILL; r += 1
        r = put(ws, r, ["店铺", "数据推荐主推品类", "在售SKU数", "数据建议动作", "★本周重点(你手填)", ""],
                fill=H_FILL, font=H_FONT, aligns=[CEN, None, CEN, None, None, None])
        top_cat, top_v = ranked[0]
        pct0 = (top_v["orders"] / total_o * 100) if total_o else 0
        r = put(ws, r, [shop, top_cat, f"{top_v['n']:.0f}",
                        f"上周{top_v['orders']:.0f}件/¥{top_v['amt']:,.0f}，占{pct0:.0f}%，建议本周上架侧重该品类新款/变体/套装",
                        "", ""],
                aligns=[CEN, None, CEN, None, None, None])
        r += 2
    if not any_data:
        ws.merge_cells(f"A{r}:F{r}")
        ws[f"A{r}"] = "全部店铺无数据：请导出 *-上周数据.xlsx（或 *-30天数据.xlsx）后重跑"
        ws[f"A{r}"].fill = R_FILL
    else:
        ws.merge_cells(f"A{r}:F{r}")
        ws[f"A{r}"] = "📌 本周上架总方向（跨店）：把上面各店『数据推荐主推品类』汇总，就是你本周上架重点；在★列写你最终决定"
        ws[f"A{r}"].font = B_FONT; ws[f"A{r}"].fill = G_FILL; r += 1
    return r


def main():
    global DATA, OUT, SNAP, LISTING_TRACK_FILE, TODAY, WEEK_START, WEEK_END
    global SALES_FILE, FT_LIST_FILE, SETTLE_FILE, LIST_FILE, FT_FILE, SHOP_FILE, CAT_FILE, CAT_WEEK_FILE
    global SALES_SHOP, KPI

    args = parse_args()
    cfg, cfg_path = load_config(args)
    if cfg is None:
        print("⚠️ 未找到 config.json / config.example.json，使用内置占位配置（不会匹配你的数据）。"
              "请复制 config.example.json 为 config.json 并填入你的店铺。")
        cfg = DEFAULT_CONFIG
    else:
        print(f"✅ 已加载配置：{cfg_path}")
    SALES_SHOP = cfg["saas_shop"]
    SHOP_FILE = cfg["shop_file"]
    CAT_FILE = cfg["cat_file"]
    CAT_WEEK_FILE = cfg["cat_week_file"]
    KPI = cfg["kpi"]

    DATA = args.data
    TODAY = datetime.date.fromisoformat(args.date) if args.date else datetime.date.today()
    OUT = args.out or os.path.join(DATA, f"速卖通_周度行动表_{TODAY.isoformat()}.xlsx")
    # 快照存脚本所在目录（用户不会动），避免每周刷新导出文件夹时丢失导致通过率一直基线
    SNAP = args.snapshot or os.path.join(os.path.dirname(os.path.abspath(__file__)), "_listing快照.json")
    LISTING_TRACK_FILE = args.track or os.path.join(DATA, "刊登数量追踪.xlsx")
    if args.week_start and args.week_end:
        WEEK_START = datetime.date.fromisoformat(args.week_start)
        WEEK_END = datetime.date.fromisoformat(args.week_end)
    else:
        WEEK_START, WEEK_END = prev_week_window(TODAY)

    # 派生文件常量（文件名约定见 SKILL.md）
    SALES_FILE = os.path.join(DATA, "五店「销售额订单数据」.xls")
    FT_LIST_FILE = os.path.join(DATA, "全托管上架数据.xlsx")
    SETTLE_FILE = os.path.join(DATA, "QTG-GLYT1 近三十天结算.xlsx")
    LIST_FILE = os.path.join(DATA, "四店「商品管理」列表导出.xls")
    FT_FILE = os.path.join(DATA, "CLYT1全托30天数据.xlsx")
    SHOP_FILE = {"YZ3": "YZ3-30天数据.xlsx", "PHCX04": "PHCX04-30天数据.xlsx",
                 "WHDS06": "WHDS06-30天数据.xlsx", "ZHHK04": "ZHHK04-30天数据.xlsx"}
    CAT_FILE = {"YZ3": "YZ3商品总数据.xlsx", "PHCX04": "PHCX04商品总数据.xlsx",
                "WHDS06": "WHDS06商品总数据.xlsx", "ZHHK04": "ZHHK04商品总数据.xlsx"}
    CAT_WEEK_FILE = {"YZ3": "YZ3-上周数据.xlsx", "PHCX04": "PHCX04-上周数据.xlsx",
                     "WHDS06": "WHDS06-上周数据.xlsx", "ZHHK04": "ZHHK04-上周数据.xlsx"}

    tracker_created = ensure_tracker(LISTING_TRACK_FILE)

    d30 = parse_30(); catalog = parse_catalog(); ft_items = parse_ft()
    by_shop, all_ids = parse_listing()
    sales_res, sales_win = parse_sales()
    listing_counts = parse_listing_tracker()
    main_img, detail, hotlink_push = build_actions(d30, catalog)
    rate, total_online, total_prior, total_new, total_rate, baseline_week = listing_rate(by_shop, all_ids)
    kpis, dead_ids = compute_kpis(d30, ft_items, by_shop, catalog, sales_res, sales_win, listing_counts)
    cat_week = parse_cat_week()

    wb = openpyxl.Workbook()

    # ---------- 页1 公司 KPI 与本周缺口 ----------
    ws = wb.active; ws.title = "1_公司KPI与本周缺口"
    setw(ws, [12, 12, 14, 14, 14, 14, 36])
    ws.merge_cells("A1:G1"); ws["A1"] = "公司 KPI 与本周缺口（9 月考核）"; ws["A1"].font = TITLE_F
    ws.merge_cells("A2:G2")
    ws["A2"] = f"生成 {TODAY} ｜ 销售额=订单文件(POP近{kpis['销售额']['win_days']}天月化+全托30天) ｜ 利润率=订单供货价估算(扣3%佣金) ｜ 其余为当前进度快照"
    ws["A2"].font = SUB_F
    r = 4
    ws.merge_cells(f"A{r}:G{r}"); ws[f"A{r}"] = "★ 本周先看这里：缺口在哪、要补多少"; ws[f"A{r}"].font = B_FONT; ws[f"A{r}"].fill = T_FILL; r += 1
    r = put(ws, r, ["KPI", "月度目标", "当前进度", "缺口/完成率", "本周应补", "权重", "判断与动作"],
            fill=H_FILL, font=H_FONT, aligns=[CEN, CEN, CEN, CEN, CEN, CEN, None])

    s = kpis["销售额"]; gap = s["target"] - s["current"]; wk_gap = s["target"] / 4 - s["current"]
    r = put(ws, r, ["销售额", "¥{:,.0f}".format(s["target"]), "¥{:,.0f}".format(s["current"]),
                f"月化缺口 ¥{gap:,.0f}\n({s['current']/s['target']*100:.1f}%)",
                (f"已超周线 ✅ (周线 ¥{s['target']/4:,.0f})" if wk_gap <= 0 else f"周目标 ¥{wk_gap:,.0f}"), "35%",
                f"构成：四店POP近{s['win_days']}天 ¥{s['pop']:,.0f}（月化）+ 全托近30天 ¥{s['ft']:,.0f}。"
                f"⚠️ POP窗口仅{s['win_days']}天、全托为30天，月化为估算；建议改导『月至今/自然月』口径直接对标"],
          fill=G_FILL if s["current"] >= s["target"] * 0.9 else R_FILL, aligns=[CEN, CEN, CEN, None, CEN, CEN, None])

    s = kpis["刊登数量"]; gap = s["target"] - s["current"]; wk_gap = s["target"] / 4
    r = put(ws, r, ["刊登数量", f"{s['target']} 个/月", f"上周上 {int(s['current'])} 个\n(POP {int(s['pop'])} / 全托 {int(s['ft'])})",
                f"缺口 {max(0,int(gap))} 个\n(上周进度{s['current']/s['target']*100:.1f}%)",
                f"本周需上 {max(0,int(wk_gap))} 个", "10%",
                f"来源：刊登数量追踪.xlsx（用户每日维护）。上周({WEEK_START}~{WEEK_END})合计{int(s['current'])}个，其中 POP 四店{int(s['pop'])}个、CLYT1全托{int(s['ft'])}个；POP日均{int(s['pop']/max(1,(WEEK_END-WEEK_START).days+1))}个、CLYT1日均{int(s['ft']/max(1,(WEEK_END-WEEK_START).days+1))}个"],
          fill=G_FILL if s["current"] >= wk_gap else R_FILL, aligns=[CEN, CEN, CEN, None, CEN, CEN, None])

    s = kpis["热销链接"]; gap = s["target"] - s["current"]
    r = put(ws, r, ["热销链接", f"+{s['target']} 个/月", f"当前 {s['current']} 个(≥15单)",
                f"缺口 {max(0,gap)} 个", "优先推近15单的品到15单", "10%",
                f"见页2「热销链接冲刺」；当前距15单差1-7单的候选 {len(hotlink_push)} 个"],
          fill=G_FILL if s["current"] >= s["target"] else R_FILL, aligns=[CEN, CEN, CEN, None, CEN, CEN, None])

    s = kpis["全托管订单"]; gap = s["target"] - s["current"]
    r = put(ws, r, ["全托管订单", f"{s['target']} 单/月", f"近30天 {s['current']} 单",
                f"缺口 {max(0,gap)} 单\n({s['current']/s['target']*100:.1f}%)",
                f"本周补 {max(0,int(gap/3))} 单", "15%",
                "CLYT1每天固定上20个链接；优先上公司内部有动销的品"],
          fill=G_FILL if s["current"] >= s["target"] * 0.85 else R_FILL, aligns=[CEN, CEN, CEN, None, CEN, CEN, None])

    s = kpis["滞销清理"]; denom = s.get("denom")
    if denom is None:
        r = put(ws, r, ["滞销清理", "—", "⚠️ 商品总数据未打开保存(空壳)", "—", "—", "10%",
                    "本周暂停：请把4个「商品总数据」双击打开→保存一次丢回目录，重跑即得滞销清单"],
              fill=R_FILL, aligns=[CEN, CEN, CEN, None, CEN, CEN, None])
    else:
        r = put(ws, r, ["滞销清理", f"清理 {s['target']} 个\n(滞销品{denom}×60%)",
                    "当前清理：待你填报", "—", "本周清理 ≥10 个", "10%",
                    f"全量目录中近30天0销量品 {denom} 个；见页3「滞销清单TOP20」，先优化再下架"],
              fill=Y_FILL, aligns=[CEN, CEN, CEN, None, CEN, CEN, None])

    s = kpis["利润率"]; st = s.get("store", {})
    st_line = "；".join(f"{k}净{st[k]['net']*100:.0f}%" for k in ("ZHHK04", "YZ3", "WHDS06", "PHCX04") if k in st)
    r = put(ws, r, ["利润率", "≥15%", f"净 {s['current']*100:.0f}%",
                f"毛 {s['gross']*100:.0f}%（扣3%佣金前）", "—", "15%",
                f"已按订单供货价估算（非待成本表）：{st_line}。⚠️ 未扣头程运费/退款/平台其他费，为商品级毛利上限；全托待结算数据"],
          fill=G_FILL if s["current"] >= 0.15 else R_FILL, aligns=[CEN, CEN, CEN, None, CEN, CEN, None])

    r += 1
    ws.merge_cells(f"A{r}:G{r}");     ws[f"A{r}"] = ("口径说明：① 销售额=POP订单文件近" + str(kpis['销售额']['win_days']) + "天月化 + 全托30天原值，月化为估算，建议改导『月至今』直接对标；"
        "② 利润率=订单供货价估算的商品级毛利（扣3%佣金），未扣头程运费/退款/其他费，为上限；③ 刊登数=刊登数量追踪.xlsx（用户每日维护），上周窗口" + str(WEEK_START) + "~" + str(WEEK_END) + "；"
        "④ 热销链接=四店+全托近30天支付件数≥15；⑤ 滞销品=全量目录中近30天0支付件数；⑥ 通过率=列表导出ID差集法(本周−上周基线)。")
    ws[f"A{r}"].font = SUB_F

    # ---------- 页2 本周行动清单 ----------
    ws2 = wb.create_sheet("2_本周行动清单")
    setw(ws2, [6, 9, 16, 26, 11, 10, 10, 12, 38])
    ws2.merge_cells("A1:I1"); ws2["A1"] = "本周行动清单（按公司 KPI 缺口 + 流量承接排序）"; ws2["A1"].font = TITLE_F
    ws2.merge_cells("A2:I2")
    ws2["A2"] = f"生成 {TODAY} ｜ 先做 KPI 缺口动作，再做产品优化；预计耗时：KPI动作>产品优化"
    ws2["A2"].font = SUB_F
    r = 4

    ws2.merge_cells(f"A{r}:I{r}"); ws2[f"A{r}"] = "★ 本周重点（按影响排序，从 #1 开始）"; ws2[f"A{r}"].font = B_FONT; ws2[f"A{r}"].fill = T_FILL; r += 1
    bullets = []
    weekly_listing_target = KPI["刊登数量"]["target"] / 4
    if kpis["刊登数量"]["current"] < weekly_listing_target:
        gap = int(weekly_listing_target - kpis["刊登数量"]["current"])
        bullets.append(f"📌 刊登补量：上周合计 {int(kpis['刊登数量']['current'])} 个（POP {int(kpis['刊登数量']['pop'])} / 全托 {int(kpis['刊登数量']['ft'])}），距周线目标 {int(weekly_listing_target)} 还差 {gap} 个。按组1/组2轮流补 POP，CLYT1 每天20个自然达成全托")
    else:
        bullets.append(f"✅ 刊登进度正常：上周合计 {int(kpis['刊登数量']['current'])} 个（POP {int(kpis['刊登数量']['pop'])} / 全托 {int(kpis['刊登数量']['ft'])}），已超周线目标 {int(weekly_listing_target)} 个")
    if kpis["全托管订单"]["current"] < KPI["全托管订单"]["target"]:
        bullets.append(f"📌 全托管订单：近30天 {int(kpis['全托管订单']['current'])} 单，目标 {KPI['全托管订单']['target']} 单；保持 CLYT1 每天20个上新，优先公司内部有动销的品")
    if kpis["销售额"]["current"] < KPI["销售额"]["target"]:
        bullets.append(f"📌 保销售额：月化估算 ¥{kpis['销售额']['current']:,.0f}（POP近{kpis['销售额']['win_days']}天¥{kpis['销售额']['pop']:,.0f}月化+全托30天¥{kpis['销售额']['ft']:,.0f}），目标 ¥{KPI['销售额']['target']:,.0f}。主推高转化/高客单款，必要时组套提价")
    if kpis["利润率"]["current"] >= 0.15:
        bullets.append(f"✅ 利润率健康：四店POP净毛利均>60%（扣3%佣金后），远超20%红线；重点从『保利润』转向『冲规模/提客单』")
    if kpis["热销链接"]["current"] >= KPI["热销链接"]["target"]:
        bullets.append(f"✅ 热销链接已达标：当前 {kpis['热销链接']['current']} 个 ≥15单，目标 {KPI['热销链接']['target']} 个；继续把近15单的 {len(hotlink_push)} 个品推上去，稳住基本盘")
    else:
        bullets.append(f"📌 冲热销链接：当前 {kpis['热销链接']['current']} 个，目标 {KPI['热销链接']['target']} 个；优先把近15单的 {len(hotlink_push)} 个品推到15单")
    nd = len(dead_ids) if dead_ids else None
    if nd:
        bullets.append(f"🗑️ 滞销清理：全量目录中近30天0销量品 {nd} 个，本周先处理 TOP10（见页3）；处理完回填数量")
    else:
        bullets.append(f"🗑️ 滞销清理：⚠️ 商品总数据未打开保存，滞销清单暂停；请打开保存后重跑")
    bullets.append(f"🎨 主图优化 TOP{TOPN}：预计挽回 ~{int(sum(x['lost'] for x in main_img[:TOPN]))} 次点击/30天")
    bullets.append(f"🎨 详情页优化 TOP{TOPN}：预计挽回 ~{int(sum(x['lost_o'] for x in detail[:TOPN]))} 单/30天")
    bullets.append(f"🧭 本周上品方向：见页4「上周好品类」——各店『数据推荐主推品类』已按上周表现列出，请在★列写你本周重点")
    for b in bullets:
        ws2.merge_cells(f"A{r}:I{r}"); ws2[f"A{r}"] = b; ws2[f"A{r}"].alignment = WRAP; ws2[f"A{r}"].fill = B_FILL; r += 1
    r += 1

    if hotlink_push:
        ws2.merge_cells(f"A{r}:I{r}"); ws2[f"A{r}"] = f"区块A ｜ 热销链接冲刺（当前出单 8-14 件，距 15 单最近，共 {len(hotlink_push)} 个，取 TOP{TOPN}）"
        ws2[f"A{r}"].font = B_FONT; ws2[f"A{r}"].fill = Y_FILL; r += 1
        r = put(ws2, r, ["优先级", "店铺", "商品ID", "品名", "当前出单", "缺口单数", "转化率", "—", "动作"],
              fill=H_FILL, font=H_FONT, aligns=[CEN, CEN, CEN, None, CEN, CEN, CEN, CEN, None])
        for i, x in enumerate(hotlink_push[:TOPN]):
            r = put(ws2, r, [i + 1, x["shop"], x["pid"], x["title"], int(x["orders"]), int(x["gap_o"]),
                        f"{x['cvr']*100:.1f}%", "", x["act"]],
                  aligns=[CEN, CEN, CEN, None, CEN, CEN, CEN, CEN, None])
        r += 1

    ws2.merge_cells(f"A{r}:I{r}"); ws2[f"A{r}"] = f"区块B ｜ 主图/搜索卡片优化（高曝光低点击，取影响最大 TOP{TOPN}）"
    ws2[f"A{r}"].font = B_FONT; ws2[f"A{r}"].fill = Y_FILL; r += 1
    r = put(ws2, r, ["优先级", "店铺", "商品ID", "品名", "曝光量", "点击率", "店均CTR", "预计挽回点击", "动作"],
          fill=H_FILL, font=H_FONT, aligns=[CEN, CEN, CEN, None, CEN, CEN, CEN, CEN, None])
    for i, x in enumerate(main_img[:TOPN]):
        r = put(ws2, r, [i + 1, x["shop"], x["pid"], x["title"], int(x["expose"]),
                    f"{x['ctr']*100:.1f}%", f"{x['m_ctr']*100:.1f}%", int(x["lost"]), x["act"]],
              aligns=[CEN, CEN, CEN, None, CEN, CEN, CEN, CEN, None])
    r += 1

    ws2.merge_cells(f"A{r}:I{r}"); ws2[f"A{r}"] = f"区块C ｜ 详情页/标题优化（有流量不成交，取影响最大 TOP{TOPN}）"
    ws2[f"A{r}"].font = B_FONT; ws2[f"A{r}"].fill = Y_FILL; r += 1
    r = put(ws2, r, ["优先级", "店铺", "商品ID", "品名", "访客数", "转化率", "店均CVR", "预计挽回订单", "动作"],
          fill=H_FILL, font=H_FONT, aligns=[CEN, CEN, CEN, None, CEN, CEN, CEN, CEN, None])
    for i, x in enumerate(detail[:TOPN]):
        r = put(ws2, r, [i + 1, x["shop"], x["pid"], x["title"], int(x["uv"]),
                    f"{x['cvr']*100:.1f}%", f"{x['m_cvr']*100:.1f}%", int(x["lost_o"]), x["act"]],
              aligns=[CEN, CEN, CEN, None, CEN, CEN, CEN, CEN, None])
    r += 1

    ws2.merge_cells(f"A{r}:I{r}"); ws2[f"A{r}"] = f"区块D ｜ 上品通过率复盘（{'基线周，已存快照' if baseline_week else f'总通过率 {total_rate*100:.0f}%'}）"
    ws2[f"A{r}"].font = B_FONT; ws2[f"A{r}"].fill = Y_FILL; r += 1
    r = put(ws2, r, ["店铺", "本周在线", "上周基线", "新增", "计划", "通过率", "说明"],
          fill=H_FILL, font=H_FONT, aligns=[CEN, CEN, CEN, CEN, CEN, CEN, None])
    for shop in SHOP_FILE:
        d = rate[shop]; rate_s = "基线周" if d["rate"] is None else f"{d['rate']*100:.0f}%"
        note = "已存基线" if d["rate"] is None else ("⚠️>100%" if d["rate"] and d["rate"] > 1 else ("⚠️差集为负,口径待确认" if d["rate"] <= 0 else "正常"))
        r = put(ws2, r, [shop, d["online"], d["prior"], d["new"], PLAN_PER_WEEK, rate_s, note],
              aligns=[CEN, CEN, CEN, CEN, CEN, CEN, None])
    r = put(ws2, r, ["合计", total_online, total_prior, total_new, PLAN_PER_WEEK,
                 "基线周" if total_rate is None else f"{total_rate*100:.0f}%", ""],
          fill=B_FILL, font=B_FONT, aligns=[CEN, CEN, CEN, CEN, CEN, CEN, None])

    # ---------- 页3 数据底稿 ----------
    ws3 = wb.create_sheet("3_数据底稿")
    setw(ws3, [9, 16, 11, 10, 10, 12, 12, 12, 10])
    ws3.merge_cells("A1:I1"); ws3["A1"] = "数据底稿（流量/转化/滞销明细，支撑页1-2）"; ws3["A1"].font = TITLE_F
    r = 3

    ws3.merge_cells(f"A{r}:I{r}"); ws3[f"A{r}"] = f"热销链接候选（≥8单且<15单，按出单升序）"; ws3[f"A{r}"].font = B_FONT; ws3[f"A{r}"].fill = B_FILL; r += 1
    r = put(ws3, r, ["商品ID", "品名", "店铺", "曝光", "点击率", "访客", "转化率", "支付件数", "动作"],
          fill=H_FILL, font=H_FONT, aligns=[CEN, None, CEN, CEN, CEN, CEN, CEN, CEN, None])
    for x in hotlink_push[:25]:
        title = catalog.get(x["pid"], {}).get("title", "")
        r = put(ws3, r, [x["pid"], title[:28], x["shop"], int(x["expose"]), f"{x['ctr']*100:.1f}%",
                     int(x["uv"]), f"{x['cvr']*100:.1f}%", int(x["orders"]), "推到15单"],
              aligns=[CEN, None, CEN, CEN, CEN, CEN, CEN, CEN, None])
    r += 1

    dead_list = []
    for pid in dead_ids:
        info = catalog.get(pid, {}); title = info.get("title", "")
        shop = "—"
        for s, items in d30.items():
            if any(x["pid"] == pid for x in items): shop = s; break
        dead_list.append({"pid": pid, "title": title, "shop": shop})
    dead_list.sort(key=lambda z: (z["shop"], z["pid"]))
    if not dead_list:
        ws3.merge_cells(f"A{r}:I{r}"); ws3[f"A{r}"] = "滞销品清单：⚠️ 商品总数据未打开保存(空壳)，滞销清单暂停；请打开保存后重跑"
        ws3[f"A{r}"].font = B_FONT; ws3[f"A{r}"].fill = R_FILL; r += 1
    else:
        ws3.merge_cells(f"A{r}:I{r}"); ws3[f"A{r}"] = f"滞销品清单 TOP20（全量目录中近30天0支付件数，共 {len(dead_list)} 个）"; ws3[f"A{r}"].font = B_FONT; ws3[f"A{r}"].fill = B_FILL; r += 1
        r = put(ws3, r, ["商品ID", "品名", "店铺", "—", "—", "—", "—", "—", "建议动作"],
              fill=H_FILL, font=H_FONT, aligns=[CEN, None, CEN, CEN, CEN, CEN, CEN, CEN, None])
        for i, x in enumerate(dead_list[:20]):
            act = "先检查是否下架/缺货；有库存则改标题词/调价/报活动；仍不动则下架"
            r = put(ws3, r, [x["pid"], x["title"][:40], x["shop"], "", "", "", "", "", act],
                  aligns=[CEN, None, CEN, CEN, CEN, CEN, CEN, CEN, None])
        r += 1

    for shop in SHOP_FILE:
        items = d30.get(shop, [])
        if not items: continue
        ws3.merge_cells(f"A{r}:I{r}"); ws3[f"A{r}"] = f"{shop} 流量 TOP15（按曝光降序）"; ws3[f"A{r}"].font = B_FONT; ws3[f"A{r}"].fill = B_FILL; r += 1
        r = put(ws3, r, ["商品ID", "品名", "曝光", "点击率", "访客", "转化率", "成交金额", "收藏", "加购"],
              fill=H_FILL, font=H_FONT, aligns=[CEN, None, CEN, CEN, CEN, CEN, CEN, CEN, CEN])
        for x in sorted(items, key=lambda z: z["expose"], reverse=True)[:15]:
            title = catalog.get(x["pid"], {}).get("title", "")
            r = put(ws3, r, [x["pid"], title[:28], int(x["expose"]), f"{x['ctr']*100:.1f}%",
                         int(x["uv"]), f"{x['cvr']*100:.1f}%", round(x["amt"], 0), int(x["fav"]), int(x["cart"])],
                  aligns=[CEN, None, CEN, CEN, CEN, CEN, CEN, CEN, CEN])
        r += 1

    build_cat_week_sheet(wb, cat_week)
    for w in wb.worksheets:
        w.freeze_panes = "A5"; w.sheet_view.showGridLines = False
    wb.save(OUT)
    print(f"OK -> {OUT}")
    print(f"[{TODAY}] 刊登窗口 {WEEK_START}~{WEEK_END} ｜ tracker{'（本次自动创建模板，请填报后重跑）' if tracker_created else '已读取'}")
    print(f"销售额(月化)={kpis['销售额']['current']:,.0f} (POP近{kpis['销售额']['win_days']}天={kpis['销售额']['pop']:,.0f}, 全托30天={kpis['销售额']['ft']:,.0f})")
    print(f"利润率净={kpis['利润率']['current']*100:.1f}% (毛={kpis['利润率']['gross']*100:.1f}%)  刊登={kpis['刊登数量']['current']}(POP{kpis['刊登数量']['pop']}/全托{kpis['刊登数量']['ft']})")
    print(f"热销={kpis['热销链接']['current']} 全托订单={kpis['全托管订单']['current']} 滞销={len(dead_ids)} 主图={len(main_img)} 详情={len(detail)} 热销冲刺={len(hotlink_push)}")
    print(f"通过率基线={'是' if baseline_week else '否'}  刊登追踪表存在={'是' if os.path.exists(LISTING_TRACK_FILE) else '否'}")
    print(f"上周品类模块：{'回退30天文件(⚠️未导出上周文件)' if CAT_WEEK_USED_FALLBACK else '已读*-上周数据.xlsx'} ｜ 见页4")


# ==================== 参数 & 辅助 ====================
def prev_week_window(report_date):
    """取 report_date 之前最近一个完整自然周（周一~周日）。"""
    last_sun = report_date - datetime.timedelta(days=(report_date.weekday() + 1) % 7)
    if last_sun >= report_date:
        last_sun -= datetime.timedelta(days=7)
    return last_sun - datetime.timedelta(days=6), last_sun


def ensure_tracker(path):
    """刊登数量追踪表不存在时，自动创建带表头的模板（含1行示例+1行说明）。返回是否新建。"""
    if os.path.exists(path):
        return False
    from openpyxl.styles import Font as _F, PatternFill as _PF, Alignment as _A
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "每日上品数"
    headers = ["日期", "CLYT1", "YZ3", "WHDS06", "PHCX04", "ZHHK04", "合计"]
    ws.append(headers)
    for h in ws[1]:
        h.font = _F(bold=True, color="FFFFFF"); h.fill = _PF("solid", fgColor="1F4E78")
        h.alignment = _A(horizontal="center")
    ws.append([f"（示例）{datetime.date.today().strftime('%Y/%m/%d')}", 20, 5, 5, 5, 5, 40])
    ws.append(["↑ 每天填上面5店当天上架数（CLYT1=全托管，其余=POP四店），删掉本说明行"])
    for col in "ABCDEFG":
        ws.column_dimensions[col].width = 14
    try:
        wb.save(path)
        return True
    except Exception as e:
        print(f"⚠️ 无法自动创建追踪表模板 {path}: {e}")
        return False


def load_config(args):
    """按优先级返回配置字典与来源路径：
       --config > <skill根>/config.json > <skill根>/config.example.json > 内置占位。
       返回 (dict, path)；都找不到时返回 (None, None)。"""
    candidates = []
    if getattr(args, "config", None):
        candidates.append(args.config)
    skill_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    candidates.append(os.path.join(skill_root, "config.json"))
    candidates.append(os.path.join(skill_root, "config.example.json"))
    for c in candidates:
        if c and os.path.exists(c):
            try:
                with open(c, encoding="utf-8") as f:
                    return json.load(f), c
            except Exception as e:
                print(f"⚠️ 读取配置失败 {c}: {e}")
    return None, None


def parse_args():
    ap = argparse.ArgumentParser(description="速卖通《周度行动表》生成器")
    ap.add_argument("--data", default="./每周导出",
                    help="每周导出数据目录（含各店30天数据/销售额订单/结算/列表导出等）")
    ap.add_argument("--config", default=None,
                    help="配置文件路径（默认 <skill根>/config.json，其次 config.example.json）")
    ap.add_argument("--out", default=None, help="输出 xlsx 路径（默认 <data>/速卖通_周度行动表_<date>.xlsx）")
    ap.add_argument("--snapshot", default=None, help="上品ID快照 json（默认 <脚本所在目录>/_listing快照.json，勿删）")
    ap.add_argument("--track", default=None, help="刊登数量追踪表 xlsx（默认 <data>/刊登数量追踪.xlsx，缺失自动建模板）")
    ap.add_argument("--date", default=None, help="报告日期 YYYY-MM-DD（默认今天）")
    ap.add_argument("--week-start", default=None, help="刊登追踪汇总窗口起始 YYYY-MM-DD（默认上一周周一）")
    ap.add_argument("--week-end", default=None, help="刊登追踪汇总窗口结束 YYYY-MM-DD（默认上一周周日）")
    return ap.parse_args()


if __name__ == "__main__":
    main()
