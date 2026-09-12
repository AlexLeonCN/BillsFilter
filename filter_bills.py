#!/usr/bin/env python3
"""Filter Alipay / WeChat / BOC bills against Budget.xlsx baseline and export import files."""

from __future__ import annotations

import csv
import re
from collections import Counter
from datetime import datetime, date, time, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pymupdf
from openpyxl import load_workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent
PERIOD_DIR = ROOT / "2026-08"
BASELINE_PATH = PERIOD_DIR / "Budget.xlsx"
TEMPLATE_PATH = ROOT / "BudgetImportTemplate.xlsx"
ALIPAY_DIR = PERIOD_DIR / "支付宝账单"
WECHAT_DIR = PERIOD_DIR / "微信账单"
BOC_DIR = PERIOD_DIR / "中国银行账单"
ALIPAY_PATH = ALIPAY_DIR / "支付宝交易明细(20250913-20260912).csv"
WECHAT_PATH = WECHAT_DIR / "微信支付账单流水文件(20250913-20260912)_20260912015624.xlsx"
ALIPAY_OUT = ALIPAY_DIR / "AlipayImport.csv"
WECHAT_OUT = WECHAT_DIR / "WechatImport.csv"
BOC_OUT = BOC_DIR / "CcbcImport.csv"
BOC_TRANS_OUT = BOC_DIR / "CcbcTrans.csv"

RECORDER = "AlexLeon"
CURRENCY = "CNY"
IMPORT_HEADER = ["分类", "子类别", "货币", "金额", "账户", "记录人", "日期", "时间", "备注"]
PLACEHOLDER_RE = re.compile(r"^[\-\—_]+$")
# Move these note keywords from CcbcImport into CcbcTrans (transfers / repayments).
BOC_TRANS_NOTE_KEYWORDS = ("银联入账", "跨行转账", "还款", "无卡交易")

# Display name used in import file for normalized semantic accounts.
SEMANTIC_DISPLAY = {
    "alex_wechat": "Alex的微信钱包",
    "ivy_wechat": "Ivy的微信钱包",
    "alex_alipay": "Alex的支付宝余额",
    "ivy_alipay": "Ivy的支付宝余额",
    "alex_huabei": "Alex的蚂蚁花呗",
    "alex_baitiao": "Alex的京东白条",
    "cash": "现金",
}

# Map raw payment-method text (after stripping discounts) to semantic keys.
SEMANTIC_ALIASES = {
    "零钱": "alex_wechat",
    "微信零钱": "alex_wechat",
    "零钱通": "alex_wechat",
    "微信钱包": "alex_wechat",
    "alex的微信钱包": "alex_wechat",
    "余额": "alex_alipay",
    "余额宝": "alex_alipay",
    "支付宝余额": "alex_alipay",
    "支付宝": "alex_alipay",
    "alex的支付宝余额": "alex_alipay",
    "花呗": "alex_huabei",
    "蚂蚁花呗": "alex_huabei",
    "alex的蚂蚁花呗": "alex_huabei",
    "京东白条": "alex_baitiao",
    "白条": "alex_baitiao",
    "alex的京东白条": "alex_baitiao",
    "现金": "cash",
}

# Known baseline card accounts: last4 -> preferred display name.
CARD_DISPLAY: dict[str, str] = {}

ALIPAY_CATEGORY_MAP = {
    "日用百货": ("购物", "日用百货"),
    "餐饮美食": ("餐饮", None),
    "交通出行": ("交通", None),
    "医疗健康": ("医疗", None),
    "保险": ("医疗", "保险"),
    "文化休闲": ("娱乐", None),
    "充值缴费": ("居家", None),
    "爱车养车": ("交通", None),
    "数码电器": ("购物", None),
    "酒店旅游": ("旅行", None),
    "服饰装扮": ("购物", None),
    "商业服务": ("购物", None),
    "公共服务": ("居家", None),
    "家居家装": ("居家", None),
    "生活服务": ("居家", None),
    "住房物业": ("居家", None),
    "母婴亲子": ("购物", None),
    "运动户外": ("运动", None),
    "美容美发": ("购物", None),
    "转账红包": ("红包", None),
    "亲友代付": ("人情", None),
    "退款": ("退款", None),
    "信用借还": ("支出平账", None),
    "其他": ("购物", None),
}

WECHAT_TYPE_MAP = {
    "商户消费": ("购物", None),
    "扫二维码付款": ("购物", None),
    "亲属卡交易": ("人情", None),
    "亲属卡交易-退款": ("退款", None),
    "转账": ("人情", None),
    "转账-退款": ("退款", None),
    "微信红包": ("红包", None),
    "微信红包（单发）": ("红包", None),
    "微信红包（群红包）": ("红包", None),
    "微信红包-退款": ("退款", None),
    "零钱提现": ("支出平账", None),
    "其他": ("购物", None),
}

CARD_LAST4_RE = re.compile(r"[（(](\d{4})[）)]")
DIGITS_RE = re.compile(r"\d{4}")


def parse_amount(value) -> Decimal | None:
    if value is None:
        return None
    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value)).quantize(Decimal("0.01"))
        except InvalidOperation:
            return None
    text = str(value).strip().replace(",", "").replace("¥", "").replace("￥", "")
    if not text or text in {"/", "-"}:
        return None
    # WeChat sometimes prefixes +/-
    text = text.replace("+", "")
    try:
        return Decimal(text).quantize(Decimal("0.01"))
    except InvalidOperation:
        return None


def parse_datetime(value) -> tuple[date | None, time | None]:
    if value is None:
        return None, None
    if isinstance(value, datetime):
        return value.date(), value.time().replace(microsecond=0)
    if isinstance(value, date) and not isinstance(value, datetime):
        return value, None
    text = str(value).strip()
    if not text:
        return None, None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M"):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.date(), dt.time()
        except ValueError:
            pass
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            d = datetime.strptime(text, fmt).date()
            return d, None
        except ValueError:
            pass
    return None, None


def format_date(d: date | None) -> str | None:
    return d.isoformat() if d else None


def format_time(t: time | None) -> str | None:
    if t is None:
        return None
    return t.strftime("%H:%M")


def strip_payment_method(raw: str | None) -> str:
    if not raw:
        return ""
    text = str(raw).strip()
    if not text or text in {"/", "-", "无"}:
        return ""
    # Alipay often appends discounts after &
    return text.split("&")[0].strip()


def clean_field(value: str | None) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", "").replace("\r", "").strip()
    if not text or PLACEHOLDER_RE.match(text):
        return ""
    return text


def extract_card_last4(text: str) -> str | None:
    if not text:
        return None
    m = CARD_LAST4_RE.search(text)
    if m:
        return m.group(1)
    # fallback: trailing 4 digits
    digits = DIGITS_RE.findall(text)
    if digits and ("银行" in text or "信用卡" in text or "储蓄卡" in text or "卡" in text):
        return digits[-1]
    return None


def bank_family(text: str) -> str | None:
    """Rough bank family token for semantic equality."""
    text = text.lower()
    mapping = [
        ("平安", "pingan"),
        ("招商", "cmb"),
        ("中国银行", "boc"),
        ("中行", "boc"),
        ("工商", "icbc"),
        ("工行", "icbc"),
        ("浦发", "spdb"),
        ("北京银行", "bob"),
        ("花呗", "huabei"),
        ("白条", "baitiao"),
        ("支付宝", "alipay"),
        ("微信", "wechat"),
        ("零钱", "wechat"),
        ("现金", "cash"),
    ]
    for needle, key in mapping:
        if needle.lower() in text:
            return key
    return None


def account_key(raw_account: str | None) -> tuple | None:
    """Return a comparable account key.

    Keys:
      ('card', last4)
      ('semantic', alias)
      ('raw', normalized_text)
    """
    primary = strip_payment_method(raw_account)
    if not primary:
        return None

    last4 = extract_card_last4(primary)
    if last4:
        return ("card", last4)

    alias = SEMANTIC_ALIASES.get(primary.lower()) or SEMANTIC_ALIASES.get(primary)
    if alias:
        return ("semantic", alias)

    # Try fuzzy semantic for names like "Alex的微信钱包" / "Ivy的支付宝余额"
    lower = primary.lower()
    if "微信" in primary or "零钱" in primary:
        owner = "ivy" if "ivy" in lower else "alex"
        return ("semantic", f"{owner}_wechat")
    if "支付宝" in primary or primary in {"余额", "余额宝"}:
        owner = "ivy" if "ivy" in lower else "alex"
        return ("semantic", f"{owner}_alipay")
    if "花呗" in primary:
        return ("semantic", "alex_huabei")
    if "白条" in primary:
        return ("semantic", "alex_baitiao")
    if "现金" in primary:
        return ("semantic", "cash")

    return ("raw", re.sub(r"\s+", "", primary).lower())


def accounts_match(key_a: tuple | None, key_b: tuple | None) -> bool:
    if key_a is None or key_b is None:
        return False
    if key_a == key_b:
        return True
    # card last4 is enough
    if key_a[0] == "card" and key_b[0] == "card":
        return key_a[1] == key_b[1]
    return False


def display_account(raw_account: str | None, key: tuple | None) -> str:
    primary = strip_payment_method(raw_account) or (raw_account or "").strip()
    if key is None:
        return primary or "未知账户"
    if key[0] == "card":
        last4 = key[1]
        if last4 in CARD_DISPLAY:
            return CARD_DISPLAY[last4]
        # Prefer cleaned primary without savings/credit wording noise if already good
        return primary or f"银行卡({last4})"
    if key[0] == "semantic":
        return SEMANTIC_DISPLAY.get(key[1], primary)
    return primary


def load_baseline_keys() -> Counter:
    wb = load_workbook(BASELINE_PATH, read_only=True, data_only=True)
    ws = wb["Records"]
    headers = next(ws.iter_rows(values_only=True))
    idx = {h: i for i, h in enumerate(headers)}

    # Also build preferred card display names from Accounts sheet if present.
    if "Accounts" in wb.sheetnames:
        aws = wb["Accounts"]
        aheaders = next(aws.iter_rows(values_only=True))
        aname_i = aheaders.index("name") if "name" in aheaders else 1
        for row in aws.iter_rows(values_only=True):
            name = row[aname_i]
            if not name:
                continue
            last4 = extract_card_last4(str(name))
            if last4 and last4 not in CARD_DISPLAY:
                CARD_DISPLAY[last4] = str(name)

    counter: Counter = Counter()
    for row in ws.iter_rows(values_only=True):
        # skip accidental header-like row
        if row[idx["账户"]] == "账户":
            continue
        d, _ = parse_datetime(row[idx["日期"]])
        if d is None and row[idx["日期"]] is not None:
            # date may already be string yyyy-mm-dd
            try:
                d = date.fromisoformat(str(row[idx["日期"]])[:10])
            except ValueError:
                d = None
        amount = parse_amount(row[idx["金额"]])
        key = account_key(row[idx["账户"]])
        if d is None or amount is None or key is None:
            continue
        # Also learn card display from records
        if key[0] == "card" and key[1] not in CARD_DISPLAY and row[idx["账户"]]:
            CARD_DISPLAY[key[1]] = str(row[idx["账户"]])
        counter[(d, key, amount)] += 1
    wb.close()
    return counter


def build_note(*parts: str | None) -> str | None:
    cleaned = []
    for p in parts:
        if p is None:
            continue
        text = str(p).strip()
        if not text or text in {"/", "-"}:
            continue
        if text not in cleaned:
            cleaned.append(text)
    return " | ".join(cleaned) if cleaned else None


def is_alipay_row_useful(row: dict) -> bool:
    status = (row.get("交易状态") or "").strip()
    if status in {"交易关闭", "已关闭", "还款失败"}:
        return False
    amount = parse_amount(row.get("金额"))
    if amount is None or amount == 0:
        return False
    # Skip pure failed empties
    pay = strip_payment_method(row.get("收/付款方式"))
    if not pay and amount == 0:
        return False
    return True


def is_wechat_row_useful(row: dict) -> bool:
    status = str(row.get("当前状态") or "").strip()
    # Keep refunds / success; skip nothing special except empty amount
    amount = parse_amount(row.get("金额(元)"))
    if amount is None or amount == 0:
        return False
    # Neutral withdraw etc. still useful if amount present
    if status in {"对方已退还"} and row.get("收/支") == "/":
        return True
    return True


def map_alipay_category(tx_cat: str | None) -> tuple[str | None, str | None]:
    if not tx_cat:
        return ("购物", None)
    return ALIPAY_CATEGORY_MAP.get(tx_cat, ("购物", None))


def map_wechat_category(tx_type: str | None, goods: str | None, peer: str | None) -> tuple[str | None, str | None]:
    text = " ".join(x for x in [tx_type, goods, peer] if x)
    if any(k in text for k in ("餐饮", "外卖", "美食", "面", "饭", "餐", "咖啡", "奶茶")):
        return ("餐饮", None)
    if any(k in text for k in ("停车", "地铁", "公交", "滴滴", "打车", "出行", "火车", "机票")):
        return ("交通", None)
    if any(k in text for k in ("医院", "药房", "医疗", "药店")):
        return ("医疗", None)
    if tx_type in WECHAT_TYPE_MAP:
        return WECHAT_TYPE_MAP[tx_type]
    if tx_type and tx_type.endswith("-退款"):
        return ("退款", None)
    return ("购物", None)


def consume_duplicate(baseline: Counter, d: date, key: tuple | None, amount: Decimal) -> bool:
    """Return True if this bill exists in baseline (and consume one occurrence)."""
    if key is None:
        return False
    # Exact key first
    exact = (d, key, amount)
    if baseline[exact] > 0:
        baseline[exact] -= 1
        return True
    # Card last4 / semantic equivalence against remaining keys with same date+amount
    for (bd, bkey, bamt), cnt in list(baseline.items()):
        if cnt <= 0:
            continue
        if bd == d and bamt == amount and accounts_match(key, bkey):
            baseline[(bd, bkey, bamt)] -= 1
            return True
    return False


def read_alipay_rows() -> list[dict]:
    rows: list[dict] = []
    with ALIPAY_PATH.open("r", encoding="gbk", newline="") as f:
        header = None
        for line in f:
            if line.startswith("交易时间"):
                header = next(csv.reader([line]))
                # trailing empty column name from Alipay export
                header = [h.strip() for h in header]
                break
        if header is None:
            raise RuntimeError("Alipay header not found")
        reader = csv.DictReader(f, fieldnames=header)
        for raw in reader:
            if not raw.get("交易时间"):
                continue
            rows.append(raw)
    return rows


def read_wechat_rows() -> list[dict]:
    wb = load_workbook(WECHAT_PATH, read_only=True, data_only=True)
    ws = wb.active
    rows: list[dict] = []
    headers = None
    for row in ws.iter_rows(values_only=True):
        if row and row[0] == "交易时间":
            headers = [str(h).strip() if h is not None else f"col{i}" for i, h in enumerate(row)]
            continue
        if headers is None:
            continue
        if row[0] is None:
            continue
        item = {headers[i]: row[i] if i < len(row) else None for i in range(len(headers))}
        rows.append(item)
    wb.close()
    return rows


def filter_alipay(baseline: Counter) -> tuple[list[dict], dict]:
    stats = Counter()
    output: list[dict] = []
    for raw in read_alipay_rows():
        stats["total"] += 1
        if not is_alipay_row_useful(raw):
            stats["skipped_useless"] += 1
            continue
        d, t = parse_datetime(raw.get("交易时间"))
        amount = parse_amount(raw.get("金额"))
        pay_raw = raw.get("收/付款方式")
        key = account_key(pay_raw)
        if d is None or amount is None:
            stats["skipped_parse"] += 1
            continue
        if key is None:
            # Still export with unknown account if amount exists? Keep but cannot match baseline.
            stats["no_account"] += 1
        if consume_duplicate(baseline, d, key, amount):
            stats["filtered_dup"] += 1
            continue

        cat, sub = map_alipay_category(raw.get("交易分类"))
        inout = (raw.get("收/支") or "").strip()
        if inout == "不计收支" and (raw.get("交易分类") or "").strip() == "退款":
            cat, sub = "退款", None
        note = build_note(
            raw.get("交易对方"),
            raw.get("商品说明"),
            raw.get("备注"),
            f"[{inout}]" if inout else None,
            f"状态:{raw.get('交易状态')}" if raw.get("交易状态") else None,
        )
        output.append(
            {
                "分类": cat,
                "子类别": sub,
                "货币": CURRENCY,
                "金额": float(amount),
                "账户": display_account(pay_raw, key),
                "记录人": RECORDER,
                "日期": format_date(d),
                "时间": format_time(t),
                "备注": note,
            }
        )
        stats["kept"] += 1
    return output, stats


def filter_wechat(baseline: Counter) -> tuple[list[dict], dict]:
    stats = Counter()
    output: list[dict] = []
    for raw in read_wechat_rows():
        stats["total"] += 1
        if not is_wechat_row_useful(raw):
            stats["skipped_useless"] += 1
            continue
        d, t = parse_datetime(raw.get("交易时间"))
        amount = parse_amount(raw.get("金额(元)"))
        pay_raw = raw.get("支付方式")
        # WeChat uses '/' when no payment method (e.g. some refunds into original channel)
        if str(pay_raw).strip() in {"/", "-", ""}:
            # try infer from 当前状态 / type
            status = str(raw.get("当前状态") or "")
            if "零钱" in status:
                pay_raw = "零钱"
            else:
                pay_raw = None
        key = account_key(pay_raw)
        if d is None or amount is None:
            stats["skipped_parse"] += 1
            continue
        if key is None:
            stats["no_account"] += 1
        if consume_duplicate(baseline, d, key, amount):
            stats["filtered_dup"] += 1
            continue

        cat, sub = map_wechat_category(raw.get("交易类型"), raw.get("商品"), raw.get("交易对方"))
        inout = str(raw.get("收/支") or "").strip()
        note = build_note(
            raw.get("交易对方"),
            raw.get("商品"),
            None if str(raw.get("备注") or "").strip() in {"", "/"} else raw.get("备注"),
            f"[{inout}]" if inout and inout != "/" else None,
            f"类型:{raw.get('交易类型')}" if raw.get("交易类型") else None,
            f"状态:{raw.get('当前状态')}" if raw.get("当前状态") else None,
        )
        output.append(
            {
                "分类": cat,
                "子类别": sub,
                "货币": CURRENCY,
                "金额": float(amount),
                "账户": display_account(pay_raw, key),
                "记录人": RECORDER,
                "日期": format_date(d),
                "时间": format_time(t),
                "备注": note,
            }
        )
        stats["kept"] += 1
    return output, stats


def write_import_xlsx(path: Path, records: list[dict]) -> None:
    template = load_workbook(TEMPLATE_PATH)
    # Keep _BudgetTemplate sheet as-is; replace Records data (remove sample).
    if "Records" in template.sheetnames:
        ws = template["Records"]
        # Clear existing data rows but keep header
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
    else:
        ws = template.active
        ws.title = "Records"

    header = IMPORT_HEADER
    # Ensure header
    for col, name in enumerate(header, 1):
        cell = ws.cell(1, col, name)
        cell.font = Font(name="Calibri", bold=True)

    for r_idx, rec in enumerate(records, 2):
        for c_idx, name in enumerate(header, 1):
            ws.cell(r_idx, c_idx, rec.get(name))

    # Refresh created_at on template meta if present
    if "_BudgetTemplate" in template.sheetnames:
        meta = template["_BudgetTemplate"]
        for row in meta.iter_rows(min_row=2, max_col=2):
            if row[0].value == "created_at_utc":
                row[1].value = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                break

    template.save(path)


def write_import_csv(path: Path, records: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=IMPORT_HEADER, extrasaction="ignore")
        writer.writeheader()
        for rec in records:
            row = {k: rec.get(k) for k in IMPORT_HEADER}
            # Keep empty cells empty rather than "None"
            for k, v in list(row.items()):
                if v is None:
                    row[k] = ""
            writer.writerow(row)


def map_boc_category(
    tx_name: str,
    memo: str,
    peer: str,
    signed_amount: Decimal,
    self_name: str | None,
) -> tuple[str | None, str | None]:
    text = f"{tx_name} {memo} {peer}"
    abs_amt = abs(signed_amount)
    is_income = signed_amount > 0

    if "结息" in tx_name:
        return ("投资", "结息")
    if "退款" in tx_name or "冲正" in tx_name:
        return ("退款", None)
    if "信用卡还款" in text:
        return ("支出平账", None)
    if "房贷" in text:
        return ("房贷", None)
    if any(k in text for k in ("房租", "月房租", "月租", "车位")):
        if "车位" in text or "沪" in text:
            return ("交通", "车位费")
        return ("居家", "住宿房租")
    if any(k in text for k in ("滴滴", "打车", "出行快车")):
        return ("交通", "打车")
    if any(k in text for k in ("地铁", "公交", "停车", "加油", "高铁", "火车")):
        return ("交通", None)
    if any(k in text for k in ("医院", "药房", "医药", "医疗", "健康")):
        return ("医疗", None)
    if any(k in text for k in ("餐饮", "外卖", "美食", "咖啡", "奶茶", "饭")):
        return ("餐饮", None)
    if tx_name in {"自助取款", "自助存款"}:
        return ("支出平账" if signed_amount < 0 else "收入平账", None)

    # Own-account transfers (counterparty is self)
    if self_name and self_name in peer:
        return ("收入平账" if is_income else "支出平账", None)

    if "跨行转账" in tx_name or "银联入账" in tx_name:
        if is_income:
            return ("人情", None) if abs_amt < 20000 else ("收入平账", None)
        if abs_amt >= 1000:
            return ("人情", None)
        return ("人情", None)

    if "提现" in tx_name:
        return ("收入平账", None)

    # Default third-party quick pay / card-not-present
    if is_income:
        return ("退款", None)
    return ("购物", None)


def read_boc_pdf(path: Path) -> tuple[str | None, str | None, list[dict]]:
    """Parse one BOC transaction PDF. Returns (card_no, customer_name, rows)."""
    doc = pymupdf.open(path)
    header_text = doc[0].get_text() if doc.page_count else ""
    card_m = re.search(r"借记卡号：(\d+)", header_text)
    name_m = re.search(r"客户姓名：([^\n\r]+)", header_text)
    card_no = card_m.group(1) if card_m else None
    customer = name_m.group(1).strip() if name_m else None

    rows: list[dict] = []
    for page_index in range(doc.page_count):
        page = doc[page_index]
        tables = page.find_tables()
        if not tables.tables:
            continue
        data = tables.tables[0].extract()
        for raw in data:
            if not raw or not str(raw[0] or "").startswith("20"):
                continue
            rows.append(
                {
                    "记账日期": raw[0],
                    "记账时间": raw[1],
                    "币别": raw[2],
                    "金额": raw[3],
                    "余额": raw[4],
                    "交易名称": clean_field(raw[5]),
                    "渠道": clean_field(raw[6]),
                    "网点名称": clean_field(raw[7]),
                    "附言": clean_field(raw[8]),
                    "对方账户名": clean_field(raw[9]),
                    "对方卡号/账号": clean_field(raw[10]),
                    "对方开户行": clean_field(raw[11]),
                    "_source_file": path.name,
                    "_card_no": card_no,
                    "_customer": customer,
                }
            )
    doc.close()
    return card_no, customer, rows


def read_boc_rows() -> list[dict]:
    if not BOC_DIR.is_dir():
        raise FileNotFoundError(f"BOC bill directory not found: {BOC_DIR}")
    all_rows: list[dict] = []
    for pdf in sorted(BOC_DIR.glob("*.pdf")):
        _card, _name, rows = read_boc_pdf(pdf)
        all_rows.extend(rows)
    return all_rows


def boc_account_raw(card_no: str | None) -> str:
    if not card_no:
        return "中国银行"
    return f"中国银行({card_no[-4:]})"


def filter_boc(baseline: Counter) -> tuple[list[dict], dict]:
    stats = Counter()
    output: list[dict] = []
    for raw in read_boc_rows():
        stats["total"] += 1
        d, t = parse_datetime(f"{raw.get('记账日期')} {raw.get('记账时间')}")
        signed = parse_amount(raw.get("金额"))
        if d is None or signed is None or signed == 0:
            stats["skipped_parse"] += 1
            continue
        amount = abs(signed)
        pay_raw = boc_account_raw(raw.get("_card_no"))
        key = account_key(pay_raw)
        if key is None:
            stats["no_account"] += 1
        if consume_duplicate(baseline, d, key, amount):
            stats["filtered_dup"] += 1
            continue

        tx_name = raw.get("交易名称") or ""
        memo = raw.get("附言") or ""
        peer = raw.get("对方账户名") or ""
        cat, sub = map_boc_category(tx_name, memo, peer, signed, raw.get("_customer"))
        inout = "收入" if signed > 0 else "支出"
        note = build_note(
            peer,
            memo,
            tx_name,
            raw.get("对方开户行"),
            f"[{inout}]",
            f"渠道:{raw.get('渠道')}" if raw.get("渠道") else None,
        )
        output.append(
            {
                "分类": cat,
                "子类别": sub,
                "货币": CURRENCY,
                "金额": float(amount),
                "账户": display_account(pay_raw, key),
                "记录人": RECORDER,
                "日期": format_date(d),
                "时间": format_time(t),
                "备注": note,
            }
        )
        stats["kept"] += 1
    return output, stats


def is_boc_transfer_note(note: str | None) -> bool:
    text = note or ""
    return any(keyword in text for keyword in BOC_TRANS_NOTE_KEYWORDS)


def split_boc_transfers(records: list[dict]) -> tuple[list[dict], list[dict]]:
    """Split BOC rows: transfers/repayments -> CcbcTrans, remainder -> CcbcImport."""
    imports: list[dict] = []
    transfers: list[dict] = []
    for rec in records:
        if is_boc_transfer_note(rec.get("备注")):
            transfers.append(rec)
        else:
            imports.append(rec)
    return imports, transfers


def main() -> None:
    print("Loading baseline from", BASELINE_PATH)
    baseline = load_baseline_keys()
    print(f"Baseline unique keys: {len(baseline)}, total entries: {sum(baseline.values())}")
    print("Card display map:", CARD_DISPLAY)

    print("\nFiltering BOC (China Bank)...")
    boc_baseline = baseline.copy()
    boc_records, boc_stats = filter_boc(boc_baseline)
    print(dict(boc_stats))

    boc_import, boc_trans = split_boc_transfers(boc_records)
    write_import_csv(BOC_OUT, boc_import)
    write_import_csv(BOC_TRANS_OUT, boc_trans)
    print(f"Wrote {BOC_OUT}: {len(boc_import)} rows")
    print(f"Wrote {BOC_TRANS_OUT}: {len(boc_trans)} rows")

    print("\nBOC import sample:")
    for r in boc_import[:5]:
        print(r)
    print("\nBOC transfer sample:")
    for r in boc_trans[:5]:
        print(r)


if __name__ == "__main__":
    main()