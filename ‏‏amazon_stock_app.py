# -*- coding: utf-8 -*-
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import time
import pandas as pd
import io
import re
import gspread.exceptions

st.set_page_config(page_title="📦 Amazon Stock | طلبات مخزون أمازون", page_icon="📦", layout="wide")

# ══ اتصال — نفس الشيت Complaints بس الصفحات بـ AMZ_ ══
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
client = gspread.authorize(creds)

def open_spreadsheet(retries=5, delay=2):
    for attempt in range(retries):
        try:
            return client.open("Complaints")
        except gspread.exceptions.APIError as e:
            if attempt < retries - 1:
                time.sleep(delay * (2 ** attempt))
            else:
                raise e

ss = open_spreadsheet()

# ══ الأوراق — كلها AMZ_ عشان متتخلطش مع نون ══
TABS_CONFIG = {
    "AMZ_Requests":    ["MSKU","ASIN","FNSKU","Quantity","Image URL","Date Added","File Name"],
    "AMZ_Approved":    ["MSKU","ASIN","FNSKU","Qty Requested","Qty Approved","Image URL","Date Added","Date Approved"],
    "AMZ_Unavailable": ["MSKU","ASIN","FNSKU","Quantity","Image URL","Date Added","Date Marked Unavailable"],
    "AMZ_Ordered":     ["MSKU","ASIN","FNSKU","Quantity","Image URL","Date Added","Order Count","Notes"],
    "AMZ_Scheduled":   ["ASN","MSKU","ASIN","FNSKU","Quantity","Schedule Date","Image URL","Date Added","Notes","Flag"],
    "AMZ_Cancelled":   ["ASN","MSKU","ASIN","FNSKU","Quantity","Schedule Date","Image URL","Date Added","Cancel Reason","Date Cancelled"],
    "AMZ_Rescheduled": ["ASN","MSKU","ASIN","FNSKU","Quantity","Old Schedule Date","Image URL","Date Added","Reschedule Reason","Date Moved"],
    "AMZ_Expired":     ["ASN","MSKU","ASIN","FNSKU","Quantity","Schedule Date","Image URL","Date Added","Date Expired"],
    "AMZ_Check":       ["ASN","MSKU","ASIN","FNSKU","Quantity","Schedule Date","Image URL","Date Added","Notes","Flag"],
    "AMZ_CancelNotif": ["ASN","MSKUs","Schedule Date","Reason","Timestamp"],
    "AMZ_Inventory":   ["ASIN","FNSKU","MSKU","Warehouse","Condition","Stock","Image URL","Date Uploaded"],
    "AMZ_Sales":       ["MSKU","ASIN","FNSKU","Title","Event Type","Fulfillment Center","Quantity","Date","Date Uploaded"],
    "AMZ_Settings":    ["Key","Value"],
}

def _init_worksheets_batch(tabs_config, extra_titles, retries=5, delay=3):
    """
    يجيب كل الأوراق الموجودة بطلب API واحد (ss.worksheets())،
    ثم ينشئ الناقصة فقط — بدل API call لكل ورقة.
    """
    for attempt in range(retries):
        try:
            existing_ws = {ws.title: ws for ws in ss.worksheets()}
            break
        except gspread.exceptions.APIError as e:
            if attempt < retries - 1:
                time.sleep(delay * (2 ** attempt))
            else:
                raise e

    result = {}

    # إنشاء الأوراق الناقصة فقط
    needed = list(tabs_config.keys()) + extra_titles
    missing_titles = [t for t in needed if t not in existing_ws]

    for title in missing_titles:
        for attempt in range(retries):
            try:
                headers = tabs_config.get(title, ["MSKU", "Image URL"])
                rows = "2000" if title in ("links m",) else "3000"
                cols = "2"    if title in ("links m",) else "15"
                ws = ss.add_worksheet(title=title, rows=rows, cols=cols)
                ws.append_row(headers)
                existing_ws[title] = ws
                time.sleep(0.3)   # انتظر بعد كل إنشاء
                break
            except gspread.exceptions.APIError as e:
                if attempt < retries - 1:
                    time.sleep(delay * (2 ** attempt))
                else:
                    raise e

    for tab in tabs_config:
        result[tab] = existing_ws[tab]

    links_ws_ref = existing_ws.get("links m")
    return result, links_ws_ref


sheets, links_ws = _init_worksheets_batch(
    TABS_CONFIG,
    extra_titles=["links m"],
)

requests_sheet     = sheets["AMZ_Requests"]
approved_sheet     = sheets["AMZ_Approved"]
unavailable_sheet  = sheets["AMZ_Unavailable"]
ordered_sheet      = sheets["AMZ_Ordered"]
scheduled_sheet    = sheets["AMZ_Scheduled"]
cancelled_sheet    = sheets["AMZ_Cancelled"]
reschedule_sheet   = sheets["AMZ_Rescheduled"]
expired_sheet      = sheets["AMZ_Expired"]
inventory_sheet    = sheets["AMZ_Inventory"]
sales_sheet        = sheets["AMZ_Sales"]
settings_sheet     = sheets["AMZ_Settings"]
cancel_notif_sheet = sheets["AMZ_CancelNotif"]

# ══ كاش — بـ prefix amz_ عشان متتخلطش مع كاش نون ══
def get_cached(sheet, force=False):
    key = f"amz_cache_{sheet.title}"
    if force or key not in st.session_state:
        st.session_state[key] = sheet.get_all_values()
    return st.session_state[key]

def clear_cache(sheet):
    key = f"amz_cache_{sheet.title}"
    if key in st.session_state:
        del st.session_state[key]

def _preload_all_caches():
    """
    يقرأ كل الصفحات مرة واحدة عند التشغيل.
    كل ورقة مش محملة → API call واحد مع 60ms بينهم لتجنب 429.
    """
    sheets_to_preload = [
        inventory_sheet, sales_sheet, settings_sheet,
        requests_sheet, approved_sheet, unavailable_sheet,
        ordered_sheet, scheduled_sheet, cancelled_sheet,
        expired_sheet, sheets["AMZ_Check"], cancel_notif_sheet,
    ]
    for sh in sheets_to_preload:
        key = f"amz_cache_{sh.title}"
        if key not in st.session_state:
            for attempt in range(4):
                try:
                    st.session_state[key] = sh.get_all_values()
                    time.sleep(0.06)   # 60ms بين كل قراءة
                    break
                except gspread.exceptions.APIError as e:
                    if "429" in str(e) or "quota" in str(e).lower():
                        time.sleep(3 * (2 ** attempt))
                    else:
                        break
                except Exception:
                    break

# ══ إعدادات ══
def load_settings():
    data = get_cached(settings_sheet)
    s = {}
    for row in data[1:]:
        if len(row) >= 2:
            s[row[0]] = row[1]
    return s

def save_setting(key, value):
    data = get_cached(settings_sheet, force=True)
    for i, row in enumerate(data[1:], start=2):
        if len(row) >= 1 and row[0] == key:
            settings_sheet.update_cell(i, 2, value)
            clear_cache(settings_sheet)
            return
    settings_sheet.append_row([key, value])
    clear_cache(settings_sheet)

def get_excluded_warehouses():
    val = load_settings().get("excluded_warehouses", "")
    if not val.strip():
        return set()
    return {w.strip().upper() for w in val.split(",") if w.strip()}

# ══ links map ══
@st.cache_data(ttl=300)
def get_links_map():
    """ترجع dict: MSKU.upper() → Image URL (من صفحة links m)"""
    data = links_ws.get_all_values()
    m = {}
    for row in data[1:]:
        if len(row) >= 2 and row[0].strip():
            m[row[0].strip().upper()] = row[1].strip()
    return m

# ══ helpers ══
def _to_int(v):
    try:
        return int(float(str(v).replace(",", "")))
    except:
        return 0

def parse_excel_date(val):
    if val is None:
        return None
    try:
        if isinstance(val, (int, float)):
            return datetime(1899, 12, 30) + timedelta(days=int(val))
        s = str(val).strip().replace(" ", "").replace("\u00a0", "")
        if "T" in s:
            s = s[:19]
            try:
                return datetime.strptime(s, "%Y-%m-%dT%H:%M:%S")
            except:
                pass
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(s[:len(fmt)], fmt)
            except:
                pass
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(s[:10], fmt)
            except:
                pass
        return None
    except:
        return None

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def file_timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def to_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

def make_empty_template(columns):
    return to_excel(pd.DataFrame(columns=columns))

def dl_btn(df, prefix, label="⬇️ Excel | Download", key=None):
    st.download_button(label, data=to_excel(df),
        file_name=f"{prefix}_{file_timestamp()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key=key or f"amz_dlbtn_{prefix}")

def parse_count_dates(cell_value):
    val = (cell_value or "").strip()
    if not val:
        return 0, ""
    m = re.match(r"^(\d+)x\s*\|\s*(.*)$", val, re.DOTALL)
    if m:
        return int(m.group(1)), m.group(2).strip()
    return 1, val

def append_count_date(rest_dates, new_count, new_date):
    rest_dates = (rest_dates or "").strip()
    if rest_dates:
        return f"{new_count}x | {rest_dates} | {new_date}"
    return f"{new_count}x | {new_date}"

# ══ Sheet helpers ══
def safe_append(sheet, row, retries=5, delay=1):
    for attempt in range(retries):
        try:
            sheet.append_row(row, value_input_option="USER_ENTERED")
            clear_cache(sheet)
            return True
        except gspread.exceptions.APIError as e:
            if "429" in str(e) or "quota" in str(e).lower():
                time.sleep(delay * (2 ** attempt))
            else:
                time.sleep(delay)
        except Exception:
            time.sleep(delay)
    return False

def safe_delete(sheet, row_idx, retries=5, delay=1):
    for attempt in range(retries):
        try:
            sheet.delete_rows(row_idx)
            clear_cache(sheet)
            return True
        except gspread.exceptions.APIError as e:
            if "429" in str(e) or "quota" in str(e).lower():
                time.sleep(delay * (2 ** attempt))
            else:
                time.sleep(delay)
        except Exception:
            time.sleep(delay)
    return False

def safe_delete_all(sheet):
    try:
        data = sheet.get_all_values()
        if len(data) > 1:
            sheet.delete_rows(2, len(data))
        clear_cache(sheet)
        return True
    except Exception:
        return False

def safe_batch_append(sheet, rows_data, retries=5, delay=1):
    if not rows_data:
        return True
    for attempt in range(retries):
        try:
            sheet.append_rows(rows_data, value_input_option="USER_ENTERED")
            clear_cache(sheet)
            return True
        except gspread.exceptions.APIError as e:
            if "429" in str(e) or "quota" in str(e).lower():
                wait = delay * (2 ** attempt)
                st.toast(f"⏳ Google Sheets API limit — waiting {wait}s...", icon="⏳")
                time.sleep(wait)
            else:
                time.sleep(delay)
        except Exception:
            time.sleep(delay)
    return False

def safe_update_row(sheet, row_idx, values, retries=4, delay=1):
    """يحدّث صف كامل بطلب API واحد بدل cell-by-cell."""
    try:
        from gspread.utils import rowcol_to_a1
        end_col_letter = rowcol_to_a1(1, len(values))[:-1]  # get col letter only
        range_str = f"A{row_idx}:{end_col_letter}{row_idx}"
    except Exception:
        range_str = f"A{row_idx}"
    for attempt in range(retries):
        try:
            sheet.update(range_str, [values])
            clear_cache(sheet)
            return True
        except gspread.exceptions.APIError as e:
            if "429" in str(e) or "quota" in str(e).lower():
                time.sleep(delay * (2 ** attempt))
            else:
                time.sleep(delay)
        except Exception:
            time.sleep(delay)
    return False

def merge_or_get_existing_row(sheet, msku):
    """يبحث عن صف بـ MSKU (العمود الأول) في الشيت."""
    data = get_cached(sheet, force=True)
    msku_up = msku.strip().upper()
    if len(data) > 1:
        for ri, row in enumerate(data[1:], start=2):
            if row and row[0].strip().upper() == msku_up:
                return ri, row
    return None, None

# ══ شبك MSKU/ASIN/FNSKU — lookup helpers ══
def _build_lookup_indexes():
    """يبني index عكسي: ASIN→MSKU و FNSKU→MSKU من inv_map."""
    asin_to_msku  = {}
    fnsku_to_msku = {}
    for msku_up, info in inv_map.items():
        a = info.get("asin","").strip().upper()
        f = info.get("fnsku","").strip().upper()
        if a: asin_to_msku[a]  = info["msku"]
        if f: fnsku_to_msku[f] = info["msku"]
    return asin_to_msku, fnsku_to_msku

def resolve_identifiers(msku_raw, asin_raw, fnsku_raw, links_map_ref):
    """
    يأخذ أي من MSKU/ASIN/FNSKU ويرجع (msku, asin, fnsku, img) مكتملة.
    لو MSKU موجود → يشبك ASIN/FNSKU/img من inv_map.
    لو ASIN أو FNSKU بس → يحوّل لـ MSKU من inv_map ثم يكمل.
    """
    asin_to_msku, fnsku_to_msku = _build_lookup_indexes()
    msku  = (msku_raw  or "").strip()
    asin  = (asin_raw  or "").strip()
    fnsku = (fnsku_raw or "").strip()
    # حاول تعرف MSKU من ASIN أو FNSKU
    if not msku:
        if asin and asin.upper() in asin_to_msku:
            msku = asin_to_msku[asin.upper()]
        elif fnsku and fnsku.upper() in fnsku_to_msku:
            msku = fnsku_to_msku[fnsku.upper()]
    # لو لقينا MSKU، شبك الباقي من inv_map
    info = inv_map.get(msku.upper(), {}) if msku else {}
    if info:
        if not asin:  asin  = info.get("asin",  "")
        if not fnsku: fnsku = info.get("fnsku", "")
    img = links_map_ref.get(msku.upper(), "") if msku else ""
    return msku, asin, fnsku, img

# ══ inv_map — المفتاح الأساسي MSKU ══
def build_inv_map(excluded_wh: set):
    inv_data = get_cached(inventory_sheet)
    inv_map = {}
    if len(inv_data) <= 1:
        return inv_map
    # AMZ_Inventory cols: ASIN(0), FNSKU(1), MSKU(2), Warehouse(3), Condition(4), Stock(5), Image URL(6), Date Uploaded(7)
    for r in inv_data[1:]:
        while len(r) < 8: r.append("")
        asin, fnsku, msku, wh, condition, stock_raw, img, date_up = \
            r[0].strip(), r[1].strip(), r[2].strip(), r[3].strip(), r[4].strip(), r[5], r[6], r[7]
        if not msku:
            continue
        msku_up = msku.upper()
        stock   = _to_int(stock_raw)
        wh_key  = f"{wh}|{condition}"
        if msku_up not in inv_map:
            inv_map[msku_up] = {
                "msku": msku, "asin": asin, "fnsku": fnsku,
                "img": img, "date": date_up, "sales": 0,
                "warehouses": {}, "total_stock": 0, "unsellable": 0,
            }
        inv_map[msku_up]["warehouses"][wh_key] = inv_map[msku_up]["warehouses"].get(wh_key, 0) + stock
        if condition.upper() == "SELLABLE" and wh.upper() not in excluded_wh:
            inv_map[msku_up]["total_stock"] += stock
        if condition.upper() == "UNSELLABLE":
            inv_map[msku_up]["unsellable"]  += stock
        if not inv_map[msku_up]["img"]   and img:   inv_map[msku_up]["img"]   = img
        if not inv_map[msku_up]["asin"]  and asin:  inv_map[msku_up]["asin"]  = asin
        if not inv_map[msku_up]["fnsku"] and fnsku: inv_map[msku_up]["fnsku"] = fnsku
    return inv_map

# ══ sales maps — المفتاح MSKU ══
def build_sales_map_monthly():
    """مبيعات آخر 30 يوم — Event Type Shipments وQty سالب — مفهرسة بـ MSKU."""
    data = get_cached(sales_sheet)
    counts = {}
    if len(data) <= 1:
        return counts
    cutoff = datetime.now() - timedelta(days=30)
    # AMZ_Sales cols: MSKU(0), ASIN(1), FNSKU(2), Title(3), Event Type(4), Fulfillment Center(5), Quantity(6), Date(7), Date Uploaded(8)
    for row in data[1:]:
        while len(row) < 8: row.append("")
        msku, event_type, qty_raw, date_str = row[0], row[4], row[6], row[7]
        if event_type.strip().lower() not in ("shipments", "shipment"):
            continue
        qty = _to_int(qty_raw)
        if qty >= 0:
            continue
        d = parse_excel_date(date_str)
        if d and d >= cutoff:
            msku_up = msku.strip().upper()
            if msku_up:
                counts[msku_up] = counts.get(msku_up, 0) + abs(qty)
    return counts

def build_daily_sales_counts(dates):
    """مبيعات يومية مفصّلة — مفهرسة بـ MSKU."""
    data = get_cached(sales_sheet)
    dates_set = set(dates)
    counts = {}
    if len(data) <= 1:
        return counts
    for row in data[1:]:
        while len(row) < 8: row.append("")
        msku, event_type, qty_raw, date_str = row[0], row[4], row[6], row[7]
        if event_type.strip().lower() not in ("shipments", "shipment"):
            continue
        qty = _to_int(qty_raw)
        if qty >= 0:
            continue
        d = parse_excel_date(date_str)
        if d and d.date() in dates_set:
            msku_up = msku.strip().upper()
            if msku_up:
                if msku_up not in counts:
                    counts[msku_up] = {dd: 0 for dd in dates}
                counts[msku_up][d.date()] += abs(qty)
    return counts

def compute_missing_inventory_rows(display_dates):
    multi_counts = build_daily_sales_counts(display_dates)
    lm = get_links_map()
    rows = []
    for msku_up, day_counts in multi_counts.items():
        if msku_up in inv_map:
            continue
        total = sum(day_counts.values())
        if total <= 0:
            continue
        est_monthly = round((total / len(display_dates)) * 30)
        rows.append({
            "msku": msku_up, "msku_up": msku_up,
            "img": lm.get(msku_up, ""),
            "day_counts": day_counts,
            "total_recent": total,
            "est_monthly_sales": est_monthly,
        })
    rows.sort(key=lambda r: -r["total_recent"])
    return rows

# ══ CancelNotifications ══
def load_cancel_notifications():
    data = get_cached(cancel_notif_sheet, force=False)
    today = datetime.now().date()
    notifs, rows_to_delete = [], []
    if len(data) <= 1:
        return notifs
    for i, row in enumerate(data[1:], start=2):
        while len(row) < 5: row.append("")
        asn, mskus_str, sdate, reason, ts = row[0], row[1], row[2], row[3], row[4]
        if not asn.strip():
            continue
        pd_ = parse_excel_date(sdate)
        if pd_ and today > pd_.date():
            rows_to_delete.append(i)
            continue
        notifs.append({
            "asn":    asn.strip(),
            "mskus":  [s.strip() for s in mskus_str.split("|") if s.strip()],
            "sdate":  sdate.strip(),
            "reason": reason.strip(),
            "ts":     ts.strip(),
        })
    for idx in sorted(rows_to_delete, reverse=True):
        safe_delete(cancel_notif_sheet, idx)
    if rows_to_delete:
        clear_cache(cancel_notif_sheet)
    return notifs

def save_cancel_notification(asn, asins_list, sdate, reason, ts):
    safe_append(cancel_notif_sheet, [asn, "|".join(asins_list), sdate, reason, ts])
    clear_cache(cancel_notif_sheet)

def delete_cancel_notification_by_asn(asn):
    data = get_cached(cancel_notif_sheet, force=True)
    for i, row in enumerate(data[1:], start=2):
        if row and row[0].strip().upper() == asn.strip().upper():
            safe_delete(cancel_notif_sheet, i)
            clear_cache(cancel_notif_sheet)
            return

def delete_all_cancel_notifications():
    safe_delete_all(cancel_notif_sheet)
    clear_cache(cancel_notif_sheet)

def check_expired_scheduled():
    data = get_cached(scheduled_sheet, force=True)
    if len(data) <= 1:
        return
    today = datetime.now().date()
    expired_rows, keep = [], []
    for i, row in enumerate(data[1:], start=2):
        while len(row) < 9: row.append("")
        d = parse_excel_date(row[5])
        if d and today > d.date():
            expired_rows.append(row[:9] + [now_str()])
        else:
            keep.append(i)
    if expired_rows:
        safe_batch_append(expired_sheet, expired_rows)
        del_idx = sorted([x for x in range(2, len(data[1:]) + 2) if x not in keep], reverse=True)
        if del_idx:
            try:
                scheduled_sheet.delete_rows(del_idx[-1], del_idx[0] - del_idx[-1] + 1)
                clear_cache(scheduled_sheet)
            except Exception:
                for idx in del_idx:
                    safe_delete(scheduled_sheet, idx)

# ══ CSS ══
st.markdown("""
<style>
.stTabs [data-baseweb="tab-list"]{gap:5px;flex-wrap:wrap;}
.stTabs [data-baseweb="tab"]{background:#1e293b;color:white;border-radius:8px;padding:6px 12px;font-weight:bold;font-size:11px;}
.stTabs [aria-selected="true"]{background:#f97316!important;}
.wh-badge{display:inline-block;border-radius:6px;padding:2px 9px;margin:2px;font-size:12px;}
.cancel-notif-card{background:linear-gradient(135deg,#2d0a0a,#1a0000);border:1px solid #ef4444;border-left:5px solid #ef4444;border-radius:10px;padding:10px 14px;margin-bottom:8px;color:white;}
.cancel-notif-card .asn-num{font-size:16px;font-weight:bold;color:#fca5a5;}
.cancel-notif-card .asin-chip{display:inline-block;background:#4b1010;color:#fca5a5;border-radius:5px;padding:1px 7px;margin:2px;font-size:11px;}
.cancel-notif-card .reason-text{color:#fcd34d;font-size:12px;}
</style>
""", unsafe_allow_html=True)

# ══ Init ══
_today_key = f"amz_expired_checked_{datetime.now().date()}"
if _today_key not in st.session_state:
    for _old in [k for k in st.session_state if k.startswith("amz_expired_checked_") and k != _today_key]:
        del st.session_state[_old]
    _preload_all_caches()   # batch read كل الصفحات مرة واحدة
    check_expired_scheduled()
    st.session_state[_today_key] = True
else:
    _preload_all_caches()   # يحمّل الصفحات اللي ناقصة بس (الموجودة في كاش تتجاهل)


if "amz_cancel_notifs_loaded" not in st.session_state:
    st.session_state["amz_cancel_notifs"] = load_cancel_notifications()
    st.session_state["amz_cancel_notifs_loaded"] = True
elif "amz_cancel_notifs" not in st.session_state:
    st.session_state["amz_cancel_notifs"] = []

excluded_wh   = get_excluded_warehouses()
inv_map       = build_inv_map(excluded_wh)
sales_monthly = build_sales_map_monthly()
for msku_up in inv_map:
    inv_map[msku_up]["sales"] = sales_monthly.get(msku_up, 0)

# ══ UI helpers ══
def show_img(img, width=75):
    if img and str(img).startswith("http"):
        st.image(img, width=width)
    else:
        st.markdown("🖼️")

def show_sku_info(msku: str):
    """عرض معلومات المنتج بناءً على MSKU (المفتاح الأساسي)."""
    info = inv_map.get(msku.strip().upper())
    if not info:
        return
    total  = info["total_stock"]
    unsell = info["unsellable"]
    sales  = info["sales"]
    asin   = info.get("asin", "")
    fnsku  = info.get("fnsku", "")
    st.markdown(
        f"📈 **مبيع شهري | Monthly Sales:** **{sales}** &nbsp;|&nbsp; "
        f"📦 **SELLABLE:** **{total}**"
        + (f" &nbsp;|&nbsp; ⚠️ **UNSELLABLE:** {unsell}" if unsell > 0 else "")
    )
    if asin or fnsku:
        st.caption(f"ASIN: `{asin}` | FNSKU: `{fnsku}`")
    badges = []
    for wh_cond, stk in sorted(info["warehouses"].items()):
        wh, cond = wh_cond.split("|", 1) if "|" in wh_cond else (wh_cond, "")
        is_ex  = wh.upper() in excluded_wh
        is_un  = cond.upper() == "UNSELLABLE"
        bg     = "#4b1010" if is_ex or is_un else "#1e3a5f"
        color  = "#fca5a5" if is_ex or is_un else "#93c5fd"
        strike = "text-decoration:line-through;" if is_ex else ""
        badges.append(f'<span class="wh-badge" style="background:{bg};color:{color};{strike}">{wh}({cond[:4]}): {stk}</span>')
    st.markdown("🏭 " + "".join(badges), unsafe_allow_html=True)

# Alias backward compat
def show_asin_info(msku: str):
    show_sku_info(msku)

def confirm_clear(key, sheet, label=""):
    if st.session_state.get(f"amz_confirm_{key}"):
        st.warning(f"⚠️ مسح كل {label}؟ | Clear all {label}?")
        cy, cn = st.columns(2)
        if cy.button("✅ نعم | Yes", key=f"amz_yes_{key}"):
            safe_delete_all(sheet)
            st.session_state[f"amz_confirm_{key}"] = False
            st.success("✅ تم المسح | Cleared")
            st.rerun()
        if cn.button("❌ لا | No", key=f"amz_no_{key}"):
            st.session_state[f"amz_confirm_{key}"] = False
            st.rerun()

def get_latest_schedule_info(msku):
    """يجيب أقرب جدولة لـ MSKU (col[1]) من Scheduled أو Check."""
    msku_up = msku.strip().upper()
    candidates = []
    for sheet_key in ("AMZ_Scheduled", "AMZ_Check"):
        data = get_cached(sheets[sheet_key])
        if len(data) <= 1:
            continue
        for row in data[1:]:
            while len(row) < 6: row.append("")
            if row[1].strip().upper() == msku_up:
                d = parse_excel_date(row[5])
                candidates.append({"asn": row[0], "date": row[5], "qty": row[4], "parsed": d, "source": sheet_key})
    if not candidates:
        return None
    dated = [c for c in candidates if c["parsed"]]
    if dated:
        dated.sort(key=lambda c: c["parsed"])
        return dated[0]
    return candidates[0]

def schedule_coverage_badge(msku, days_to_stockout, delay_days):
    sched = get_latest_schedule_info(msku)
    if not sched:
        return ("🔴 محتاج جدولة الآن | Needs scheduling now", "#ef4444", None)
    if not sched["parsed"]:
        return (f"⚠️ مجدول (ASN {sched['asn']}) بدون تاريخ واضح", "#f59e0b", sched)
    arrival     = sched["parsed"] + timedelta(days=delay_days)
    stockout_dt = datetime.now() + timedelta(days=days_to_stockout) if days_to_stockout > 0 else datetime.now()
    src_label   = "تشييك | Check" if sched["source"] == "AMZ_Check" else "الجدولة | Scheduled"
    if arrival.date() <= stockout_dt.date():
        return (f"✅ مجدول (ASN {sched['asn']}) بتاريخ {sched['date']} [{src_label}] — هيوصل قبل النفاد", "#22c55e", sched)
    else:
        return (f"🔴 مجدول (ASN {sched['asn']}) بتاريخ {sched['date']} [{src_label}] — متأخر عن موعد النفاد", "#ef4444", sched)

def get_unavailable_ordered_note(msku):
    msku_up = msku.strip().upper()
    notes = []
    data_un = get_cached(unavailable_sheet)
    if len(data_un) > 1:
        for row in data_un[1:]:
            if row and row[0].strip().upper() == msku_up:
                while len(row) < 7: row.append("")
                cnt, dates = parse_count_dates(row[6])
                notes.append(f"❌ غير متوفر سابقاً | Was unavailable ({cnt}x) — {dates}")
                break
    data_ord = get_cached(ordered_sheet)
    if len(data_ord) > 1:
        for row in data_ord[1:]:
            if row and row[0].strip().upper() == msku_up:
                while len(row) < 8: row.append("")
                cnt, dates = parse_count_dates(row[7])
                notes.append(f"🛒 تم طلبه سابقاً | Was ordered ({cnt}x) — {dates}")
                break
    return notes

def get_recent_expired_info(msku, days_back=4):
    msku_up = msku.strip().upper()
    data = get_cached(expired_sheet)
    if len(data) <= 1:
        return None
    cutoff = datetime.now().date() - timedelta(days=days_back)
    candidates = []
    # AMZ_Expired: ASN(0), MSKU(1), ASIN(2), FNSKU(3), ...
    for row in data[1:]:
        while len(row) < 10: row.append("")
        if row[1].strip().upper() != msku_up:
            continue
        d_exp = parse_excel_date(row[9])
        if d_exp and d_exp.date() >= cutoff:
            candidates.append({"asn": row[0], "schedule_date": row[5], "date_expired": row[9], "parsed_expired": d_exp})
    if not candidates:
        return None
    candidates.sort(key=lambda c: c["parsed_expired"], reverse=True)
    return candidates[0]

def render_recent_expired_note(msku, days_back=4):
    info = get_recent_expired_info(msku, days_back)
    if not info:
        return
    st.markdown(
        f'<span style="background:#7c2d12;color:#fed7aa;border-radius:6px;padding:3px 10px;font-size:12px;">'
        f'📋 كانت مجدولة (ASN {info["asn"]}) بتاريخ {info["schedule_date"]} وانتهت بتاريخ {info["date_expired"]} | Was scheduled but expired</span>',
        unsafe_allow_html=True)

def render_day_counts_md(day_counts, dates, labels):
    parts = [f"**{lbl}:** {day_counts.get(d,0)}" for d, lbl in zip(dates, labels)]
    return " &nbsp;|&nbsp; ".join(parts)

ordinal_map = {1:"الثانية|Second",2:"الثالثة|Third",3:"الرابعة|Fourth",4:"الخامسة|Fifth"}

# ══ Sidebar — إشعارات الكنسل ══
def render_sidebar_notifications():
    notifs = st.session_state.get("amz_cancel_notifs", [])
    if not notifs:
        return
    with st.sidebar:
        st.markdown("## 🔔 إشعارات الكنسل | Cancel Alerts")
        st.markdown(f"**{len(notifs)} إشعار نشط | Active Alerts**")
        st.markdown("---")
        lm = get_links_map()
        for ni, notif in enumerate(notifs):
            asns  = notif.get("mskus", [])
            chips = "".join(f'<span class="asin-chip">{a[:12]}</span>' for a in asns[:5])
            if len(asns) > 5:
                chips += f'<span class="asin-chip">+{len(asns)-5}</span>'
            st.markdown(f"""
<div class="cancel-notif-card">
  <div>🚫 <span class="asn-num">ASN: {notif.get('asn','')}</span></div>
  <div style="font-size:12px;color:#94a3b8;">📅 {notif.get('sdate','')}</div>
  <div style="margin:6px 0;">{chips}</div>
  <div class="reason-text">📝 {notif.get('reason','') or '—'}</div>
  <div style="font-size:10px;color:#64748b;margin-top:4px;">🕐 {notif.get('ts','')}</div>
</div>""", unsafe_allow_html=True)
            img_cols = st.columns(min(len(asns[:4]), 4))
            for ci2, a in enumerate(asns[:4]):
                img_url = lm.get(a.strip().upper(), "")
                with img_cols[ci2]:
                    if img_url and img_url.startswith("http"):
                        st.image(img_url, width=55, caption=a[:10])
                    else:
                        st.markdown(f"🖼️ `{a[:10]}`")
            if st.button(f"✖️ حذف #{ni+1}", key=f"amz_sb_rm_{ni}", use_container_width=True):
                delete_cancel_notification_by_asn(notif.get("asn",""))
                st.session_state["amz_cancel_notifs"].pop(ni)
                st.rerun()
            st.markdown("---")
        if st.button("🗑️ مسح كل الإشعارات | Clear All", key="amz_sb_clear_all",
                     use_container_width=True, type="secondary"):
            delete_all_cancel_notifications()
            st.session_state["amz_cancel_notifs"] = []
            st.rerun()

render_sidebar_notifications()

# ══ حساب المرحلين من تاب المبيعات ══
def compute_transferred_from_sales():
    if not inv_map:
        return []
    s = load_settings()
    sales_days = int(s.get("sales_display_days","7") or 7)
    delay_days = int(s.get("schedule_delay_days","3") or 3)
    cov_days   = int(s.get("schedule_coverage_days","15") or 15)
    today_now  = datetime.now().date()
    dates_now  = [today_now - timedelta(days=i) for i in range(1, sales_days + 1)]
    counts_now = build_daily_sales_counts(dates_now)
    result = []
    for msku_up, info in inv_map.items():
        stock      = info.get("total_stock", 0)
        sales_m    = info.get("sales", 0)
        day_counts = counts_now.get(msku_up, {d: 0 for d in dates_now})
        total_rec  = sum(day_counts.values())
        avg_daily  = (total_rec / sales_days) if sales_days > 0 else (sales_m / 30 if sales_m > 0 else 0)
        eff_avg    = avg_daily if avg_daily > 0 else (sales_m / 30 if sales_m > 0 else 0)
        days_so    = round(stock / eff_avg) if eff_avg > 0 else 9999
        if days_so >= cov_days and eff_avg > 0:
            continue
        badge_text, _, sched = schedule_coverage_badge(info["msku"], days_so, delay_days)
        un_notes = get_unavailable_ordered_note(info["msku"])
        if "محتاج جدولة" in badge_text and not sched and not un_notes:
            result.append({
                "msku": info["msku"], "msku_up": msku_up,
                "stock": stock, "sales_month": sales_m, "img": info["img"],
                "effective_avg": eff_avg, "days_to_stockout": days_so,
                "day_counts": day_counts,
            })
    return result

if "amz_transferred_skus" not in st.session_state:
    st.session_state["amz_transferred_skus"] = compute_transferred_from_sales()

# ══════════════════════════════════════════════
st.title("📦 Amazon Stock Requests | طلبات مخزون أمازون")

tabs = st.tabs([
    "📋 الطلبات | Requests",
    "✅ الموافقة | Approved",
    "❌ غير متوفر | Unavailable",
    "🛒 تم الطلب | Ordered",
    "📅 الجدولة | Scheduled",
    "☑️ تشييك | Check",
    "🚫 ملغية | Cancelled",
    "🔄 تعديل موعد | Rescheduled",
    "⚠️ تنبيهات | Alerts",
    "📊 المخزون | Inventory",
    "🔴 مراجعة المخزون | Stock Review",
    "🗂️ منتهية | Expired",
    "⚙️ الإعدادات | Settings",
    "📈 مراجعة المبيعات | Sales Review",
    "🛒 المبيعات | Sales",
    "🗓️ تحليل الجدولة | Schedule Analysis",
    "📦 مخزون بدون بيع | No Sales",
])
(tab1,tab2,tab3,tab4,tab5,tab_check,tab6,tab7,tab8,tab9,tab10,tab11,tab12,tab13,tab14,tab15,tab16) = tabs

# ══ TAB 1 — الطلبات ══
with tab1:
    st.subheader("➕ إضافة طلبات | Add Requests")
    links_map = get_links_map()
    col_m, col_t = st.columns([3,1])
    with col_t:
        st.download_button("⬇️ Template فارغ | Empty Template",
            data=make_empty_template(["MSKU","ASIN","FNSKU","Quantity"]),
            file_name=f"amz_request_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with col_m:
        method = st.radio("طريقة الإضافة | Add Method:", ["📂 رفع ملف | Upload","✏️ لصق | Paste"], horizontal=True)

    st.caption("💡 يكفي عمود واحد: **MSKU** أو **ASIN** أو **FNSKU** — الباقي بيتشبك تلقائياً من المخزون")

    added_rows, file_name_label = [], ""
    if "Upload" in method:
        uploaded = st.file_uploader("ارفع Excel أو CSV | Upload", type=["xlsx","xls","csv"], key="amz_req_upload")
        if uploaded:
            file_name_label = uploaded.name
            try:
                df_up = pd.read_csv(uploaded,dtype=str).fillna("") if uploaded.name.endswith(".csv") \
                    else pd.read_excel(uploaded,dtype=str).fillna("")
                msku_col=asin_col=fnsku_col=qty_col=None
                for c in df_up.columns:
                    cl = c.strip().lower()
                    if cl in ("msku","seller-sku","seller sku"):  msku_col  = c
                    if cl == "asin":                              asin_col  = c
                    if cl in ("fnsku","fulfillment-channel-sku"): fnsku_col = c
                    if cl in ("quantity","qty","كمية"):           qty_col   = c

                st.info(f"📊 {len(df_up)} صف | MSKU:`{msku_col}` ASIN:`{asin_col}` FNSKU:`{fnsku_col}` Qty:`{qty_col}`")
                st.dataframe(df_up.head(10), use_container_width=True, height=150)

                resolved_preview = []
                for _, row in df_up.iterrows():
                    msku_raw  = str(row[msku_col]).strip()  if msku_col  and str(row[msku_col]).strip()  not in ("","nan") else ""
                    asin_raw  = str(row[asin_col]).strip()  if asin_col  and str(row[asin_col]).strip()  not in ("","nan") else ""
                    fnsku_raw = str(row[fnsku_col]).strip() if fnsku_col and str(row[fnsku_col]).strip() not in ("","nan") else ""
                    qty       = str(row[qty_col]).strip()   if qty_col   and str(row[qty_col]).strip()   not in ("","nan") else ""
                    if not any([msku_raw, asin_raw, fnsku_raw]):
                        continue
                    msku, asin, fnsku, img = resolve_identifiers(msku_raw, asin_raw, fnsku_raw, links_map)
                    final_msku = msku or msku_raw or asin_raw or fnsku_raw
                    added_rows.append((final_msku, asin, fnsku, qty, img))
                    resolved_preview.append({"MSKU":final_msku,"ASIN":asin,"FNSKU":fnsku,"Qty":qty,"Img✓":"✅" if img else "—"})

                if resolved_preview:
                    st.markdown("**🔗 نتيجة الشبك | Resolved:**")
                    st.dataframe(pd.DataFrame(resolved_preview), use_container_width=True, height=120, hide_index=True)

            except Exception as e:
                st.error(f"❌ {e}")
    else:
        st.caption("اكتب MSKU أو ASIN أو FNSKU — واحد في كل سطر (مع الكمية اختياري)")
        pasted = st.text_area("الصق هنا | Paste here:", height=110,
            placeholder="MSKU,Qty\n75-DMSV-3LFW,5\nB0BH1F3JHV,3\nX002AMJX6V,2")
        file_name_label = "Manual Entry"
        if pasted.strip():
            resolved_preview = []
            for line in pasted.strip().splitlines():
                parts = [p.strip() for p in line.split(",")]
                raw   = parts[0] if parts else ""
                qty   = parts[1] if len(parts) > 1 else ""
                if not raw or raw.lower() in ("msku","asin","fnsku",""): continue
                raw_up = raw.upper()
                if raw_up.startswith("B0") and len(raw_up) == 10:
                    msku, asin, fnsku, img = resolve_identifiers("", raw, "", links_map)
                elif raw_up.startswith("X00") and len(raw_up) >= 10:
                    msku, asin, fnsku, img = resolve_identifiers("", "", raw, links_map)
                else:
                    msku, asin, fnsku, img = resolve_identifiers(raw, "", "", links_map)
                final_msku = msku or raw
                added_rows.append((final_msku, asin, fnsku, qty, img))
                resolved_preview.append({"MSKU":final_msku,"ASIN":asin,"FNSKU":fnsku,"Qty":qty,"Img✓":"✅" if img else "—"})
            if resolved_preview:
                st.markdown("**🔗 نتيجة الشبك | Resolved:**")
                st.dataframe(pd.DataFrame(resolved_preview), use_container_width=True, height=120, hide_index=True)
                st.success(f"✅ {len(added_rows)} صف جاهز | rows ready")

    if added_rows:
        if st.button("📤 إضافة | Add", type="primary"):
            dn = now_str()
            if safe_batch_append(requests_sheet, [[m,a,fn,q,img,dn,file_name_label] for m,a,fn,q,img in added_rows]):
                st.success(f"✅ أُضيف {len(added_rows)} صف | rows added")
                st.rerun()

    st.divider()
    st.subheader("📋 الطلبات الحالية | Current Requests")
    data=get_cached(requests_sheet)
    if len(data)<=1:
        st.info("لا توجد طلبات | No requests yet.")
    else:
        rows=data[1:]
        df_req=pd.DataFrame(rows,columns=data[0])
        c1,c2,c3,c4=st.columns(4)
        with c1: dl_btn(df_req,"amz_requests")
        with c2:
            if st.button("✅ موافقة الكل | Approve All",use_container_width=True):
                st.session_state["amz_confirm_approve_all"]=True
        with c3:
            if st.button("❌ رفض الكل | Reject All",use_container_width=True):
                st.session_state["amz_confirm_reject_all"]=True
        with c4:
            if st.button("🗑️ مسح الكل | Clear All",type="secondary",use_container_width=True):
                st.session_state["amz_confirm_clear_req"]=True

        if st.session_state.get("amz_confirm_approve_all"):
            st.warning("⚠️ موافقة على كل الطلبات؟ | Approve all?")
            cy,cn=st.columns(2)
            if cy.button("✅ نعم",key="amz_yes_app_all"):
                dn=now_str()
                safe_batch_append(approved_sheet,
                    [[r[0],r[1],r[2],r[3],r[3],r[4] if len(r)>4 else "",r[5] if len(r)>5 else "",dn] for r in rows])
                safe_delete_all(requests_sheet)
                st.session_state["amz_confirm_approve_all"]=False; st.rerun()
            if cn.button("❌ لا",key="amz_no_app_all"):
                st.session_state["amz_confirm_approve_all"]=False; st.rerun()

        if st.session_state.get("amz_confirm_reject_all"):
            st.warning("⚠️ رفض كل الطلبات؟ | Reject all?")
            cy,cn=st.columns(2)
            if cy.button("✅ نعم",key="amz_yes_rej_all"):
                dn=now_str()
                safe_batch_append(unavailable_sheet,
                    [[r[0],r[1],r[2],r[3],r[4] if len(r)>4 else "",r[5] if len(r)>5 else "",append_count_date("",1,dn)] for r in rows])
                safe_delete_all(requests_sheet)
                st.session_state["amz_confirm_reject_all"]=False; st.rerun()
            if cn.button("❌ لا",key="amz_no_rej_all"):
                st.session_state["amz_confirm_reject_all"]=False; st.rerun()

        confirm_clear("clear_req",requests_sheet,"الطلبات | Requests")

        ordered_data=get_cached(ordered_sheet)
        ordered_mskus={}
        if len(ordered_data)>1:
            for r in ordered_data[1:]:
                while len(r)<8: r.append("")
                ordered_mskus[r[0].strip().upper()]=_to_int(r[6]) if r[6] else 1

        st.write(f"**الإجمالي | Total: {len(rows)}**")
        for i,row in enumerate(rows,start=2):
            while len(row)<7: row.append("")
            msku,asin,fnsku,qty,img,date_added,fname=row[0],row[1],row[2],row[3],row[4],row[5],row[6]
            c_img,c_info,c_act=st.columns([1,4,3])
            with c_img: show_img(img,75)
            with c_info:
                st.markdown(f"**MSKU:** `{msku}`")
                show_sku_info(msku)
                st.markdown(f"**طلب | Requested Qty:** {qty}")
                st.caption(f"📅 {date_added} | 📁 {fname}")
                prev_count=ordered_mskus.get(msku.upper(),0)
                if prev_count>0:
                    ordn=ordinal_map.get(prev_count,f"{prev_count+1}")
                    st.warning(f"🔁 تم الطلب للمرة {ordn} | Already ordered {prev_count} time(s)")
            with c_act:
                ca,cb,cc,cd=st.columns(4)
                with ca:
                    with st.popover("✅ وافق\nApprove"):
                        nq=st.text_input("Approved Qty",value=qty,key=f"amz_aqty_{i}")
                        if st.button("✅ تأكيد",key=f"amz_aconf_{i}"):
                            safe_append(approved_sheet,[msku,asin,fnsku,qty,nq,img,date_added,now_str()])
                            safe_delete(requests_sheet,i); st.rerun()
                with cb:
                    if st.button("❌ غير\nمتوفر",key=f"amz_unavail_{i}"):
                        dn=now_str()
                        un_ri,un_row=merge_or_get_existing_row(unavailable_sheet,msku)
                        if un_ri:
                            while len(un_row)<7: un_row.append("")
                            cur_count,rest=parse_count_dates(un_row[6])
                            merged=append_count_date(rest,cur_count+1,dn)
                            safe_update_row(unavailable_sheet,un_ri,[un_row[0],un_row[1],un_row[2],qty,un_row[4] or img,un_row[5],merged])
                        else:
                            safe_append(unavailable_sheet,[msku,asin,fnsku,qty,img,date_added,append_count_date("",1,dn)])
                        safe_delete(requests_sheet,i); st.rerun()
                with cc:
                    if st.button("🛒 طلب\nOrder",key=f"amz_order_{i}"):
                        dn=now_str()
                        ord_ri,ord_row=merge_or_get_existing_row(ordered_sheet,msku)
                        if ord_ri:
                            while len(ord_row)<8: ord_row.append("")
                            cur_count,rest=parse_count_dates(ord_row[7])
                            new_count=cur_count+1
                            merged=append_count_date(rest,new_count,dn)
                            safe_update_row(ordered_sheet,ord_ri,[ord_row[0],ord_row[1],ord_row[2],qty,ord_row[4] or img,dn,str(new_count),merged])
                        else:
                            safe_append(ordered_sheet,[msku,asin,fnsku,qty,img,dn,"1",append_count_date("",1,dn)])
                        safe_delete(requests_sheet,i); st.rerun()
                with cd:
                    if st.button("🗑️ حذف",key=f"amz_del_req_{i}"):
                        safe_delete(requests_sheet,i); st.rerun()
            st.divider()

# ══ TAB 2 — الموافقة ══
with tab2:
    st.subheader("✅ الطلبات الموافق عليها | Approved Requests")
    data_ap=get_cached(approved_sheet)
    if len(data_ap)<=1:
        st.info("لا توجد موافقات | No approvals yet.")
    else:
        rows_ap=data_ap[1:]
        srch=st.text_input("🔍 بحث MSKU | Search MSKU",key="amz_srch_ap")
        indexed_ap=[(i+2,r) for i,r in enumerate(rows_ap)]
        filtered=[(ri,r) for ri,r in indexed_ap if not srch or srch.strip().upper() in r[0].upper()]
        df_ap=pd.DataFrame(rows_ap,columns=data_ap[0])
        c1,c2=st.columns(2)
        with c1: dl_btn(df_ap,"amz_approved")
        with c2:
            if st.button("🗑️ مسح الكل",type="secondary",key="amz_btn_clear_ap",use_container_width=True):
                st.session_state["amz_confirm_clear_ap"]=True
        confirm_clear("clear_ap",approved_sheet,"الموافقة | Approved")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_ap)}**")
        for ri,row in filtered:
            while len(row)<8: row.append("")
            msku,asin,fnsku,qty_r,qty_a,img,da,dap=row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7]
            c_img,c_info,c_del=st.columns([1,5,1])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**MSKU:** `{msku}`")
                show_sku_info(msku)
                if qty_a and qty_a!=qty_r:
                    st.markdown(f"**طلبت | Req:** {qty_r} → **وافقوا | App:** ⚠️ **{qty_a}**")
                else:
                    st.markdown(f"**Quantity:** {qty_a}")
                st.caption(f"📅 Requested: {da} | ✅ Approved: {dap}")
            with c_del:
                if st.button("🗑️",key=f"amz_del_ap_{ri}"):
                    safe_delete(approved_sheet,ri); st.rerun()
            st.divider()

# ══ TAB 3 — غير متوفر ══
with tab3:
    st.subheader("❌ غير متوفر | Unavailable")
    data_un=get_cached(unavailable_sheet)
    if len(data_un)<=1:
        st.info("لا يوجد | Nothing unavailable yet.")
    else:
        rows_un=data_un[1:]
        srch=st.text_input("🔍 بحث MSKU",key="amz_srch_un")
        indexed_un=[(i+2,r) for i,r in enumerate(rows_un)]
        filtered=[(ri,r) for ri,r in indexed_un if not srch or srch.strip().upper() in r[0].upper()]
        df_un=pd.DataFrame(rows_un,columns=data_un[0])
        c1,c2=st.columns(2)
        with c1: dl_btn(df_un,"amz_unavailable")
        with c2:
            if st.button("🗑️ مسح الكل",type="secondary",key="amz_btn_clear_un",use_container_width=True):
                st.session_state["amz_confirm_clear_un"]=True
        confirm_clear("clear_un",unavailable_sheet,"غير المتوفر | Unavailable")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_un)}**")
        for ri,row in filtered:
            while len(row)<7: row.append("")
            msku,asin,fnsku,qty,img,da,dm=row[0],row[1],row[2],row[3],row[4],row[5],row[6]
            cnt_un,dates_un=parse_count_dates(dm)
            c_img,c_info,c_act=st.columns([1,4,2])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**MSKU:** `{msku}`")
                if asin: st.caption(f"ASIN: `{asin}`")
                show_sku_info(msku)
                st.markdown(f"**Qty:** {qty}")
                if cnt_un>1: st.warning(f"🔁 تكرر {cnt_un} مرة | Marked unavailable {cnt_un}x")
                if dates_un: st.caption(f"❌ تواريخ: {dates_un}")
            with c_act:
                with st.popover("↩️ رجّع للموافقة\nReturn to Approved"):
                    nq_un=st.text_input("الكمية المعدّلة",value=qty,key=f"amz_un_ret_qty_{ri}")
                    if st.button("✅ أرسل للموافقة",key=f"amz_un_ret_conf_{ri}"):
                        safe_append(approved_sheet,[msku,asin,fnsku,qty,nq_un,img,da,now_str()])
                        safe_delete(unavailable_sheet,ri); st.rerun()
                if st.button("🗑️",key=f"amz_del_un_{ri}"):
                    safe_delete(unavailable_sheet,ri); st.rerun()
            st.divider()

# ══ TAB 4 — تم الطلب ══
with tab4:
    st.subheader("🛒 تم الطلب | Ordered Items")
    data_ord=get_cached(ordered_sheet)
    if len(data_ord)<=1:
        st.info("لا يوجد | No ordered items yet.")
    else:
        rows_ord=data_ord[1:]
        srch=st.text_input("🔍 بحث MSKU",key="amz_srch_ord")
        indexed_ord=[(i+2,r) for i,r in enumerate(rows_ord)]
        filtered=[(ri,r) for ri,r in indexed_ord if not srch or srch.strip().upper() in r[0].upper()]
        df_ord=pd.DataFrame(rows_ord,columns=data_ord[0])
        c1,c2=st.columns(2)
        with c1: dl_btn(df_ord,"amz_ordered")
        with c2:
            if st.button("🗑️ مسح الكل",type="secondary",key="amz_btn_clear_ord",use_container_width=True):
                st.session_state["amz_confirm_clear_ord"]=True
        confirm_clear("clear_ord",ordered_sheet,"تم الطلب | Ordered")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_ord)}**")
        for ri,row in filtered:
            while len(row)<8: row.append("")
            msku,asin,fnsku,qty,img,da,cnt,note=row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7]
            cnt_ord,dates_ord=parse_count_dates(note)
            c_img,c_info,c_act=st.columns([1,4,2])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**MSKU:** `{msku}`")
                if asin: st.caption(f"ASIN: `{asin}`")
                show_sku_info(msku)
                st.markdown(f"**Quantity:** {qty}")
                if cnt_ord>1: st.warning(f"🔁 تكرر {cnt_ord} مرة | Ordered {cnt_ord}x")
                if dates_ord: st.caption(f"🗓️ تواريخ الطلب: {dates_ord}")
                st.caption(f"📅 آخر تحديث: {da} | 🔢 عدد الطلبات: {cnt}")
            with c_act:
                ca,cb=st.columns(2)
                with ca:
                    with st.popover("↩️ رجّع\nReturn"):
                        nq=st.text_input("الكمية المعدّلة",value=qty,key=f"amz_ret_qty_{ri}")
                        if st.button("✅ أرسل للموافقة",key=f"amz_ret_conf_{ri}"):
                            safe_append(approved_sheet,[msku,asin,fnsku,qty,nq,img,da,now_str()])
                            safe_delete(ordered_sheet,ri); st.rerun()
                with cb:
                    if st.button("🗑️",key=f"amz_del_ord_{ri}"):
                        safe_delete(ordered_sheet,ri); st.rerun()
            st.divider()

# ══ TAB 5 — الجدولة ══
with tab5:
    st.subheader("📅 الجدولة | Scheduled Items")
    links_map=get_links_map()
    col_t,_=st.columns([1,3])
    with col_t:
        st.download_button("⬇️ Template الجدولة",
            data=make_empty_template(["ASN","MSKU","ASIN","FNSKU","Qty","Schedule Date"]),
            file_name=f"amz_schedule_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    st.caption("💡 يكفي عمود واحد: **MSKU** أو **ASIN** أو **FNSKU** — الباقي يتشبك تلقائياً")

    upl_sc=st.file_uploader("ارفع ملف الجدولة | Upload Schedule File",type=["xlsx","xls","csv"],key="amz_sched_upload")
    if upl_sc:
        try:
            df_sc=pd.read_csv(upl_sc,dtype=str).fillna("") if upl_sc.name.endswith(".csv") \
                else pd.read_excel(upl_sc,dtype=str).fillna("")
            cm={}
            for c in df_sc.columns:
                cl=c.strip().lower()
                if cl=="asn":                                         cm["asn"]  =c
                if cl in ("msku","seller-sku","seller sku"):          cm["msku"] =c
                if cl=="asin":                                        cm["asin"] =c
                if cl in ("fnsku","fulfillment-channel-sku"):         cm["fnsku"]=c
                if cl in ("qty","quantity","كمية"):                   cm["qty"]  =c
                if "جدول" in cl or "schedule" in cl or cl=="date":   cm["date"] =c
            asn_c  = cm.get("asn")
            msku_c = cm.get("msku")
            asin_c = cm.get("asin")
            fnsk_c = cm.get("fnsku")
            qty_c  = cm.get("qty")
            date_c = cm.get("date")

            st.info(f"📊 {len(df_sc)} صف | ASN:`{asn_c}` MSKU:`{msku_c}` ASIN:`{asin_c}` FNSKU:`{fnsk_c}` Qty:`{qty_c}` Date:`{date_c}`")
            st.dataframe(df_sc, use_container_width=True, height=150)

            # تاريخ افتراضي لو الملف مفيهوش
            col_da, col_db = st.columns([2,3])
            with col_da:
                fallback_date = st.date_input(
                    "📅 تاريخ الجدولة الافتراضي (لو مفيش تاريخ في الملف)",
                    value=datetime.now().date() + timedelta(days=3),
                    key="amz_sched_fallback_date")
            with col_db:
                st.markdown(""); st.markdown("")
                st.caption("← يُطبَّق فقط على الصفوف التي لا تحتوي تاريخاً")

            if st.button("📤 إضافة الجدولة | Add Schedule", type="primary"):
                existing=get_cached(scheduled_sheet, force=True)
                ex_pairs=set()
                if len(existing)>1:
                    for r in existing[1:]:
                        while len(r)<2: r.append("")
                        ex_pairs.add((r[0].strip().upper(), r[1].strip().upper()))
                dn=now_str(); to_add=[]; skipped=0; resolved_log=[]
                for _, row in df_sc.iterrows():
                    asn     = str(row[asn_c]).strip()    if asn_c  and str(row[asn_c]).strip()   not in ("","nan") else ""
                    msku_raw= str(row[msku_c]).strip()   if msku_c and str(row[msku_c]).strip()  not in ("","nan") else ""
                    asin_raw= str(row[asin_c]).strip()   if asin_c and str(row[asin_c]).strip()  not in ("","nan") else ""
                    fnsk_raw= str(row[fnsk_c]).strip()   if fnsk_c and str(row[fnsk_c]).strip()  not in ("","nan") else ""
                    qty     = str(row[qty_c]).strip()    if qty_c  and str(row[qty_c]).strip()   not in ("","nan") else ""
                    dval    = str(row[date_c]).strip()   if date_c and str(row[date_c]).strip()  not in ("","nan") else ""
                    if not any([msku_raw, asin_raw, fnsk_raw]):
                        continue
                    msku, asin, fnsku, _ = resolve_identifiers(msku_raw, asin_raw, fnsk_raw, links_map)

                    # الصورة دائماً من الـ MSKU النهائي
                    img = links_map.get(msku.strip().upper(), "") if msku else ""
                    final_msku = msku or msku_raw or asin_raw or fnsk_raw
                    pd_ = parse_excel_date(dval)
                    ds  = pd_.strftime("%Y-%m-%d") if pd_ else fallback_date.strftime("%Y-%m-%d")
                    pair = (asn.upper(), final_msku.upper())
                    if pair in ex_pairs:
                        skipped += 1
                    else:
                        to_add.append([asn, final_msku, asin, fnsku, qty, ds, img, dn, "", ""])
                        if asn: ex_pairs.add(pair)
                    resolved_log.append({"ASN":asn,"MSKU":final_msku,"ASIN":asin,"FNSKU":fnsku,"Qty":qty,"Date":ds,"Img✓":"✅" if img else "—"})
                safe_batch_append(scheduled_sheet, to_add)
                msg=f"✅ أُضيف | Added: {len(to_add)}"
                if skipped: msg+=f" | ⚠️ مكرر: {skipped}"
                st.success(msg)
                if resolved_log:
                    st.dataframe(pd.DataFrame(resolved_log[:20]), use_container_width=True, height=150, hide_index=True)
                st.rerun()
        except Exception as e:
            st.error(f"❌ {e}")

    # ── إضافة يدوية سريعة ──
    with st.expander("✏️ إضافة جدولة يدوية | Quick Manual Entry", expanded=False):
        st.caption("أدخل MSKU أو ASIN أو FNSKU — واحد يكفي")
        mc1,mc2,mc3 = st.columns(3)
        with mc1: man_id   = st.text_input("MSKU / ASIN / FNSKU", key="amz_man_id", placeholder="75-DMSV-3LFW")
        with mc2: man_qty  = st.text_input("الكمية | Qty", key="amz_man_qty", placeholder="10")
        with mc3: man_date = st.date_input("تاريخ الجدولة", value=datetime.now().date()+timedelta(days=3), key="amz_man_date")
        man_asn = st.text_input("ASN (اختياري)", key="amz_man_asn", placeholder="FBA123456")
        if st.button("➕ أضف | Add", key="amz_man_add", type="secondary"):
            if man_id.strip():
                rid = man_id.strip().upper()
                if rid.startswith("B0") and len(rid)==10:
                    msku_m,asin_m,fnsku_m,img_m = resolve_identifiers("",man_id.strip(),"",links_map)
                elif rid.startswith("X00") and len(rid)>=10:
                    msku_m,asin_m,fnsku_m,img_m = resolve_identifiers("","",man_id.strip(),links_map)
                else:
                    msku_m,asin_m,fnsku_m,img_m = resolve_identifiers(man_id.strip(),"","",links_map)
                final_m = msku_m or man_id.strip()
                safe_append(scheduled_sheet,[man_asn.strip(),final_m,asin_m,fnsku_m,man_qty.strip(),
                    man_date.strftime("%Y-%m-%d"),img_m,now_str(),"",""])
                st.success(f"✅ أُضيف: MSKU={final_m} | Date={man_date.strftime('%Y-%m-%d')}")
                st.rerun()
            else:
                st.warning("أدخل MSKU أو ASIN أو FNSKU")

    st.divider()
    st.subheader("📋 الجدولة الحالية | Current Schedule")
    data_sch=get_cached(scheduled_sheet)
    if len(data_sch)<=1:
        st.info("لا توجد جدولة | No scheduled items.")
    else:
        rows_sch=data_sch[1:]
        def sort_key_sch(r):
            d=parse_excel_date(r[5] if len(r)>5 else "")
            return d if d else datetime(2099,1,1)
        rows_sch_sorted=sorted(rows_sch,key=sort_key_sch)

        cancel_notif_asns={n["asn"].upper() for n in st.session_state.get("amz_cancel_notifs",[])}
        chk_data_t5=get_cached(sheets["AMZ_Check"])
        checked_asns=set()
        if len(chk_data_t5)>1:
            for cr in chk_data_t5[1:]:
                if cr: checked_asns.add(cr[0].strip().upper())

        asn_groups={}
        for r in rows_sch_sorted:
            while len(r)<10: r.append("")
            asn=r[0].strip()
            if asn not in asn_groups:
                asn_groups[asn]={"date":r[5],"items":[],"checked":asn.upper() in checked_asns}
            asn_groups[asn]["items"].append(r)

        df_sch=pd.DataFrame(rows_sch,columns=data_sch[0])
        c1,c2=st.columns(2)
        with c1: dl_btn(df_sch,"amz_scheduled")
        with c2:
            if st.button("🗑️ مسح الكل",type="secondary",key="amz_btn_clear_sc",use_container_width=True):
                st.session_state["amz_confirm_clear_sc"]=True
        confirm_clear("clear_sc",scheduled_sheet,"الجدولة | Schedule")

        c_s1,c_s2=st.columns(2)
        with c_s1: srch_asn=st.text_input("🔍 بحث ASN",key="amz_srch_asn")
        with c_s2: srch_asin_sch=st.text_input("🔍 بحث MSKU",key="amz_srch_msku_sch")
        today=datetime.now().date()
        st.write(f"**إجمالي ASN | Total ASNs: {len(asn_groups)}**")

        for asn,group in asn_groups.items():
            if srch_asn and srch_asn.strip().upper() not in asn.upper(): continue
            items_=group["items"]
            if srch_asin_sch and not any(srch_asin_sch.strip().upper() in r[1].strip().upper() for r in items_): continue
            sdate=group["date"]
            pd_date=parse_excel_date(sdate)
            is_exp=pd_date and today>pd_date.date()
            has_cancel_notif=asn.upper() in cancel_notif_asns
            has_alert=any(
                inv_map.get(r[1].strip().upper(),{}).get("sales",0)>0 and
                _to_int(r[4])>inv_map.get(r[1].strip().upper(),{}).get("sales",0)
                for r in items_)
            border="#ef4444" if has_alert else "#f59e0b" if is_exp else "#3b82f6"
            bg="#2d1515" if has_alert else "#2d2000" if is_exp else "#0f172a"
            cancel_badge=(' &nbsp;<span style="background:#7f1d1d;color:#fca5a5;border-radius:6px;padding:2px 10px;font-size:12px;font-weight:bold;">🚫 اتشيك واتكنسل</span>') if has_cancel_notif else ""
            st.markdown(
                f'<div style="border-left:5px solid {border};background:{bg};color:white;border-radius:10px;padding:8px 14px;margin-bottom:4px;">'
                f'<b>ASN:</b> {asn} &nbsp;|&nbsp; 📅 <b>Schedule Date:</b> <b>{sdate}</b>{cancel_badge}</div>',
                unsafe_allow_html=True)

            if has_cancel_notif:
                for notif in st.session_state.get("amz_cancel_notifs",[]):
                    if notif.get("asn","").upper()==asn.upper():
                        st.markdown(
                            f'<div style="background:#1a0000;border:1px solid #ef4444;border-radius:8px;padding:8px 12px;margin:4px 0 8px 0;">'
                            f'<span style="color:#fca5a5;font-weight:bold;">🚫 تم الكنسل من التشييك | Cancelled from Check</span><br>'
                            f'<span style="color:#fcd34d;font-size:12px;">📝 السبب: {notif.get("reason","") or "—"}</span><br>'
                            f'<span style="color:#94a3b8;font-size:11px;">🕐 {notif.get("ts","")}</span></div>',
                            unsafe_allow_html=True)
                        lm_t5=get_links_map()
                        notif_mskus=notif.get("mskus",[])
                        if notif_mskus:
                            ic=st.columns(min(len(notif_mskus[:6]),6))
                            for ci3,a3 in enumerate(notif_mskus[:6]):
                                iu3=lm_t5.get(a3.strip().upper(),"")
                                with ic[ci3]:
                                    if iu3 and iu3.startswith("http"): st.image(iu3,width=60,caption=a3[:10])
                                    else: st.markdown(f"🖼️ `{a3[:10]}`")
                        break

            for r in items_:
                while len(r)<10: r.append("")
                msku_r,qty,img=r[1].strip(),r[4],r[6]
                info=inv_map.get(msku_r.upper(),{})
                monthly=info.get("sales",0)
                is_al=monthly>0 and _to_int(qty)>monthly
                c_img2,c_info2=st.columns([1,6])
                with c_img2: show_img(img,60)
                with c_info2:
                    st.markdown(f"&nbsp;&nbsp;**MSKU:** `{msku_r}` | **Qty:** {qty}")
                    show_sku_info(msku_r)
                    if is_al: st.markdown(f"&nbsp;&nbsp;🔴 **تنبيه:** الكمية ({qty}) > المبيع الشهري ({monthly})")

            ca,cb,cc,cd=st.columns(4)
            with ca:
                with st.popover("☑️ Check"):
                    select_all=st.checkbox("تحديد الكل",key=f"amz_chk_all_{asn}")
                    selected_mskus={}
                    for ri2,r in enumerate(items_):
                        while len(r)<10: r.append("")
                        msku2=r[1].strip()
                        selected_mskus[msku2]=st.checkbox(f"`{msku2}` — Qty:{r[4]}",value=select_all,key=f"amz_chk_msku_{asn}_{ri2}")
                    if st.button("✅ أرسل للتشييك",key=f"amz_send_chk_{asn}"):
                        dn=now_str(); all_sel=all(selected_mskus.values())
                        to_add=[]
                        for r in items_:
                            msku2=r[1].strip()
                            flag="" if all_sel else ("highlighted" if selected_mskus.get(msku2,False) else "")
                            to_add.append([r[0],r[1],r[2],r[3],r[4],r[5],r[6],dn,"",flag])
                        safe_batch_append(sheets["AMZ_Check"],to_add)
                        sch_d=get_cached(scheduled_sheet,force=True)
                        del_i=[i2 for i2,sr in enumerate(sch_d[1:],start=2) if sr[0].strip().upper()==asn.upper()]
                        for i2 in sorted(del_i,reverse=True): safe_delete(scheduled_sheet,i2)
                        st.success(f"☑️ أُرسل للتشييك — ASN: {asn}"); st.rerun()
            with cb:
                with st.popover("🚫 كنسل - غير متوفر\nCancel"):
                    reason_u=st.text_input("سبب | Reason",key=f"amz_rsn_u_{asn}")
                    if st.button("✅ تأكيد",key=f"amz_can_u_{asn}"):
                        dn=now_str()
                        to_add=[[r[0],r[1],r[2],r[3],r[4],r[5],r[6],r[7],f"غير متوفر — {reason_u}",dn] for r in items_]
                        safe_batch_append(cancelled_sheet,to_add)
                        sch_d=get_cached(scheduled_sheet,force=True)
                        del_idx=[idx for idx,sr in enumerate(sch_d[1:],start=2) if sr[0].strip().upper()==asn.upper()]
                        for idx in sorted(del_idx,reverse=True): safe_delete(scheduled_sheet,idx)
                        st.success("🚫 تم الكنسل"); st.rerun()
            with cc:
                with st.popover("🔄 تغيير موعد\nReschedule"):
                    reason_r=st.text_input("سبب التغيير",key=f"amz_rsn_r_{asn}")
                    if st.button("✅ تأكيد",key=f"amz_can_r_{asn}"):
                        dn=now_str()
                        to_add=[[r[0],r[1],r[2],r[3],r[4],r[5],r[6],r[7],reason_r,dn] for r in items_]
                        safe_batch_append(reschedule_sheet,to_add)
                        sch_d=get_cached(scheduled_sheet,force=True)
                        del_idx=[idx for idx,sr in enumerate(sch_d[1:],start=2) if sr[0].strip().upper()==asn.upper()]
                        for idx in sorted(del_idx,reverse=True): safe_delete(scheduled_sheet,idx)
                        st.success("🔄 تم التعديل"); st.rerun()
            with cd:
                status="⚠️ منتهي | Expired" if is_exp else "✅ ساري | Active"
                st.markdown(f"&nbsp;{status}")
            st.divider()

# ══ TAB CHECK ══
with tab_check:
    st.subheader("☑️ قيد التشييك | Under Check")
    if st.session_state.get("amz_cancel_notifs"):
        st.markdown("---")
        st.markdown("### 🔔 إشعارات الكنسل الأخيرة")
        for notif in st.session_state["amz_cancel_notifs"]:
            mskus_str2=", ".join(notif.get("mskus",[])[:5])
            st.error(f"🚫 ASN **{notif.get('asn','')}** (📅 {notif.get('sdate','')}) — MSKUs: {mskus_str2} — السبب: {notif.get('reason','')} — {notif.get('ts','')}")
        if st.button("✖️ مسح الإشعارات",key="amz_clear_notifs"):
            delete_all_cancel_notifications()
            st.session_state["amz_cancel_notifs"]=[]
            st.rerun()
        st.markdown("---")

    data_chk=get_cached(sheets["AMZ_Check"])
    if len(data_chk)<=1:
        st.info("لا يوجد | No items under check.")
    else:
        rows_chk=data_chk[1:]
        chk_groups={}
        for idx,r in enumerate(rows_chk,start=2):
            while len(r)<10: r.append("")
            asn=r[0].strip()
            if asn not in chk_groups:
                chk_groups[asn]={"date":r[5],"items":[],"indices":[]}
            chk_groups[asn]["items"].append(r)
            chk_groups[asn]["indices"].append(idx)

        df_chk=pd.DataFrame(rows_chk,columns=data_chk[0])
        c1,c2=st.columns(2)
        with c1: dl_btn(df_chk,"amz_check")
        with c2:
            if st.button("🗑️ مسح الكل",type="secondary",key="amz_btn_clear_chk",use_container_width=True):
                st.session_state["amz_confirm_clear_chk"]=True
        confirm_clear("clear_chk",sheets["AMZ_Check"],"التشييك | Check")
        st.write(f"**إجمالي ASN: {len(chk_groups)}**")

        for asn,grp in chk_groups.items():
            sdate=grp["date"]; items_=grp["items"]
            has_hl=any(len(r)>9 and r[9]=="highlighted" for r in items_)
            st.markdown(
                f'<div style="border-left:5px solid #8b5cf6;background:#1a0a2e;border-radius:10px;padding:8px 14px;margin-bottom:4px;color:white;">'
                f'<b>ASN:</b> {asn} &nbsp;|&nbsp; 📅 <b>Schedule Date:</b> <b>{sdate}</b>'
                +(' &nbsp; 🔴 <b>يوجد MSKUs مميزة</b>' if has_hl else '')
                +'</div>',unsafe_allow_html=True)

            for r in items_:
                while len(r)<10: r.append("")
                msku_r,qty,img,flag=r[1].strip(),r[4],r[6],r[9]
                is_hl=flag=="highlighted"
                c_img2,c_info2=st.columns([1,6])
                with c_img2: show_img(img,60)
                with c_info2:
                    tag=" 🔴 **مميز | Highlighted**" if is_hl else ""
                    st.markdown(f"**MSKU:** `{msku_r}` | **Qty:** {qty}{tag}")
                    show_sku_info(msku_r)

            ca,cb=st.columns(2)
            with ca:
                if st.button(f"↩️ رجّع للجدولة — {asn}",key=f"amz_ret_chk_{asn}",type="primary"):
                    dn=now_str(); lm=get_links_map()
                    to_add=[[r[0],r[1],r[2],r[3],r[4],r[5],lm.get(r[1].strip().upper(),r[6]),dn,"تم تشييكه",""] for r in items_]
                    safe_batch_append(scheduled_sheet,to_add)
                    for idx in sorted(grp["indices"],reverse=True): safe_delete(sheets["AMZ_Check"],idx)
                    st.success(f"✅ تم الإرجاع — ASN: {asn}"); st.rerun()
            with cb:
                with st.popover(f"🚫 كنسل — {asn}"):
                    cancel_reason=st.text_input("سبب الكنسل",key=f"amz_chk_rsn_{asn}")
                    if st.button("✅ تأكيد الكنسل",key=f"amz_chk_can_{asn}"):
                        dn=now_str()
                        to_add=[[r[0],r[1],r[2],r[3],r[4],r[5],r[6],r[7],f"تشييك — {cancel_reason}",dn] for r in items_]
                        safe_batch_append(cancelled_sheet,to_add)
                        for idx in sorted(grp["indices"],reverse=True): safe_delete(sheets["AMZ_Check"],idx)
                        hl_mskus=[r[1].strip() for r in items_ if len(r)>9 and r[9]=="highlighted"]
                        all_mskus=[r[1].strip() for r in items_]
                        notif_mskus=hl_mskus if hl_mskus else all_mskus
                        new_notif={"asn":asn,"sdate":sdate,"mskus":notif_mskus,"reason":cancel_reason,"ts":dn}
                        save_cancel_notification(asn,notif_mskus,sdate,cancel_reason,dn)
                        if "amz_cancel_notifs" not in st.session_state:
                            st.session_state["amz_cancel_notifs"]=[]
                        st.session_state["amz_cancel_notifs"].insert(0,new_notif)
                        st.session_state["amz_cancel_notifs"]=st.session_state["amz_cancel_notifs"][:50]
                        st.success("🚫 تم الكنسل"); st.rerun()
            st.divider()

# ══ TAB 6 — ملغية ══
with tab6:
    st.subheader("🚫 الجدولة الملغية | Cancelled Schedule")
    data_can=get_cached(cancelled_sheet)
    if len(data_can)<=1:
        st.info("لا يوجد | No cancelled schedules.")
    else:
        rows_can=data_can[1:]
        srch=st.text_input("🔍 بحث ASN",key="amz_srch_can")
        indexed_can=[(i+2,r) for i,r in enumerate(rows_can)]
        filtered=[(ri,r) for ri,r in indexed_can if not srch or srch.strip().upper() in r[0].upper()]
        df_can=pd.DataFrame(rows_can,columns=data_can[0])
        c1,c2=st.columns(2)
        with c1: dl_btn(df_can,"amz_cancelled")
        with c2:
            if st.button("🗑️ مسح الكل",type="secondary",key="amz_btn_clear_can",use_container_width=True):
                st.session_state["amz_confirm_clear_can"]=True
        confirm_clear("clear_can",cancelled_sheet,"الملغية | Cancelled")
        st.write(f"**عرض | Showing: {len(filtered)} / {len(rows_can)}**")
        for ri,row in filtered:
            while len(row)<10: row.append("")
            asn,msku_r,asin_r,fnsku,qty,sd,img,dadd,reason,dcan=\
                row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9]
            c_img,c_info,c_del=st.columns([1,5,1])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**ASN:** `{asn}` | **MSKU:** `{msku_r}`")
                if asin_r: st.caption(f"ASIN: `{asin_r}`")
                show_sku_info(msku_r)
                st.markdown(f"**Qty:** {qty}")
                st.caption(f"📅 Schedule: {sd} | 🚫 Cancelled: {dcan}")
                if reason: st.caption(f"📝 السبب: {reason}")
            with c_del:
                if st.button("🗑️",key=f"amz_del_can_{ri}"):
                    safe_delete(cancelled_sheet,ri); st.rerun()
            st.divider()

# ══ TAB 7 — تعديل الموعد ══
with tab7:
    st.subheader("🔄 تعديل الموعد | Rescheduled Items")
    data_res=get_cached(reschedule_sheet)
    if len(data_res)<=1:
        st.info("لا يوجد | No rescheduled items.")
    else:
        rows_res=data_res[1:]
        asn_res_groups={}
        for idx,r in enumerate(rows_res,start=2):
            while len(r)<10: r.append("")
            asn=r[0].strip()
            if asn not in asn_res_groups:
                asn_res_groups[asn]={"old_date":r[5],"reason":r[8],"items":[],"indices":[]}
            asn_res_groups[asn]["items"].append(r)
            asn_res_groups[asn]["indices"].append(idx)

        df_res=pd.DataFrame(rows_res,columns=data_res[0])
        c1,c2=st.columns(2)
        with c1: dl_btn(df_res,"amz_rescheduled")
        with c2:
            if st.button("🗑️ مسح الكل",type="secondary",key="amz_btn_clear_res",use_container_width=True):
                st.session_state["amz_confirm_clear_res"]=True
        confirm_clear("clear_res",reschedule_sheet,"تعديل الموعد | Rescheduled")

        links_map2=get_links_map()
        for asn,grp in asn_res_groups.items():
            st.markdown(
                f'<div style="border-left:5px solid #f59e0b;background:#1a1500;border-radius:10px;padding:8px 14px;margin-bottom:4px;color:white;">'
                f'<span style="font-size:15px;font-weight:bold;">ASN: {asn}</span><br>'
                f'<span>📅 <b style="color:#fcd34d;">الموعد القديم | Old Date: {grp["old_date"]}</b></span></div>',
                unsafe_allow_html=True)
            if grp["reason"]: st.caption(f"📝 سبب التعديل: {grp['reason']}")
            with st.expander(f"✏️ تعديل وإرجاع للجدولة — ASN {asn}",expanded=False):
                new_asn=st.text_input("ASN جديد | New ASN",value=asn,key=f"amz_new_asn_{asn}")
                new_date=st.text_input("تاريخ جديد (YYYY-MM-DD)",value="",key=f"amz_new_date_{asn}",placeholder="2025-08-15")
                edited_items=[]
                for ri2,r in enumerate(grp["items"]):
                    while len(r)<8: r.append("")
                    msku_r,asin_r,fnsk,qty,img=r[1].strip(),r[2],r[3],r[4],r[6]
                    c_img2,c_s2,c_q2=st.columns([1,3,2])
                    with c_img2: show_img(img,55)
                    with c_s2:
                        st.markdown(f"**MSKU:** `{msku_r}`")
                        if asin_r: st.caption(f"ASIN: `{asin_r}`")
                        show_sku_info(msku_r)
                    with c_q2:
                        new_qty=st.text_input("Qty",value=qty,key=f"amz_res_qty_{asn}_{ri2}")
                    edited_items.append((msku_r,asin_r,fnsk,new_qty,img))
                if st.button("✅ أرجع للجدولة",key=f"amz_ret_sch_{asn}",type="primary"):
                    if not new_date.strip():
                        st.error("❌ أدخل تاريخ جديد")
                    else:
                        dn=now_str()
                        to_add=[[new_asn,msku_r,asin_r,fnsk,qty,new_date,links_map2.get(msku_r.upper(),img),dn,"",""]
                                for msku_r,asin_r,fnsk,qty,img in edited_items]
                        safe_batch_append(scheduled_sheet,to_add)
                        for idx in sorted(grp["indices"],reverse=True): safe_delete(reschedule_sheet,idx)
                        st.success(f"✅ تم الإرجاع للجدولة — ASN: {new_asn}"); st.rerun()
            st.divider()

# ══ TAB 8 — تنبيهات ══
with tab8:
    st.subheader("⚠️ تنبيهات الجدولة | Schedule Alerts")
    st.caption("الكمية المجدولة أعلى من المبيع الشهري | Scheduled qty > Monthly sales")
    data_sc8=get_cached(scheduled_sheet)
    alerts=[]
    if len(data_sc8)>1:
        for row in data_sc8[1:]:
            while len(row)<8: row.append("")
            asn,msku_r,asin_r,fnsku,qty,sdate,img=row[0],row[1],row[2],row[3],row[4],row[5],row[6]
            info=inv_map.get(msku_r.upper(),{})
            monthly=info.get("sales",0)
            stock=info.get("total_stock",0)
            try:
                if monthly>0 and _to_int(qty)>monthly:
                    alerts.append((asn,msku_r,qty,monthly,stock,sdate,img))
            except: pass
    if not inv_map:
        st.info("ارفع ملف المخزون أولاً | Upload Inventory first")
    elif not alerts:
        st.success("✅ لا توجد تنبيهات | No alerts")
    else:
        df_al=pd.DataFrame(alerts,columns=["ASN","MSKU","Scheduled Qty","Monthly Sales","Total Stock","Schedule Date","Image URL"])
        c1,c2=st.columns(2)
        with c1: dl_btn(df_al,"amz_alerts")
        with c2: st.error(f"⚠️ تنبيهات | Alerts: {len(alerts)}")
        for asn,msku_r,qty,monthly,stock,sdate,img in alerts:
            c_img,c_info=st.columns([1,6])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**ASN:** `{asn}` | **MSKU:** `{msku_r}`")
                show_sku_info(msku_r)
                st.markdown(f"🔴 **الكمية المجدولة:** {qty} > **المبيع الشهري:** {monthly}")
                st.caption(f"📅 تاريخ الجدولة: {sdate}")
            st.divider()

# ══ TAB 9 — المخزون ══
with tab9:
    st.subheader("📊 المخزون | Inventory")
    links_map=get_links_map()
    col_t,_=st.columns([1,3])
    with col_t:
        st.download_button("⬇️ Template المخزون",
            data=make_empty_template(["ASIN","FNSKU","MSKU","Warehouse","Condition","Stock"]),
            file_name=f"amz_inventory_template_{file_timestamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    st.caption("يقبل تقرير FBA Inventory من Amazon — أعمدة: asin, fulfillment-channel-sku (FNSKU), seller-sku (MSKU), Warehouse-Condition-code, Quantity Available")
    upl_inv=st.file_uploader("ارفع ملف المخزون | Upload Inventory File",type=["xlsx","xls","xlsm","csv"],key="amz_inv_upload")
    if upl_inv:
        try:
            df_inv_up=pd.read_csv(upl_inv,dtype=str).fillna("") if upl_inv.name.endswith(".csv") \
                else pd.read_excel(upl_inv,dtype=str).fillna("")
            asin_c=fnsku_c=msku_c=wh_c=cond_c=stock_c=None
            for c in df_inv_up.columns:
                cl=c.strip().lower()
                if cl=="asin": asin_c=c
                if cl in ("fnsku","fulfillment-channel-sku"): fnsku_c=c
                if cl in ("msku","seller-sku"): msku_c=c
                if "warehouse" in cl: wh_c=c
                if "condition" in cl:
                    cond_c = c

                if cl == "stock":
                    stock_c = c

                elif "quantity" in cl and "available" in cl:
                    stock_c = c

                elif "quantity" in cl and not stock_c:
                    stock_c = c
                
                elif "quantity" in cl and not stock_c: stock_c=c
            if not asin_c:
                for c in df_inv_up.columns:
                    if "asin" in c.lower(): asin_c=c; break
            if not fnsku_c:
                for c in df_inv_up.columns:
                    if "fnsku" in c.lower() or "fulfillment-channel" in c.lower(): fnsku_c=c; break
            if not msku_c:
                for c in df_inv_up.columns:
                    if "seller-sku" in c.lower() or "msku" in c.lower(): msku_c=c; break
            if not cond_c:
                for c in df_inv_up.columns:
                    if "condition" in c.lower(): cond_c=c; break
            if not stock_c:
               for c in df_inv_up.columns:
                   cl = c.strip().lower()

                   if (
                       cl == "stock"
                       or "quantity" in cl
                       or "available" in cl
                       or "qty" in cl
                   ):
                       stock_c = c
                       break
            st.info(f"📊 {len(df_inv_up)} صف | ASIN:`{asin_c}` FNSKU:`{fnsku_c}` MSKU:`{msku_c}` Cond:`{cond_c}` Stock:`{stock_c}` WH:`{wh_c}`")
            st.dataframe(df_inv_up.head(10),use_container_width=True,height=150)

            def do_inv_upload(replace=False):
                dn=now_str(); to_add=[]
                for _,row in df_inv_up.iterrows():
                    asin  =str(row[asin_c]).strip()  if asin_c  else ""
                    fnsku =str(row[fnsku_c]).strip() if fnsku_c else ""
                    msku  =str(row[msku_c]).strip()  if msku_c  else ""
                    wh    =str(row[wh_c]).strip()    if wh_c    else "FBA"
                    cond  =str(row[cond_c]).strip()  if cond_c  else "SELLABLE"
                    stk   =str(row[stock_c]).strip() if stock_c else "0"
                    img   =links_map.get(msku.upper(),"")
                    if asin and asin.lower()!="nan":
                        to_add.append([asin,fnsku,msku,wh,cond,stk,img,dn])
                if replace: safe_delete_all(inventory_sheet)
                safe_batch_append(inventory_sheet,to_add)
                clear_cache(inventory_sheet)
                return len(to_add)

            ca,cb=st.columns(2)
            with ca:
                if st.button("📤 إضافة للموجود | Append",type="primary",use_container_width=True):
                    n=do_inv_upload(replace=False)
                    st.success(f"✅ أُضيف {n} صف"); st.rerun()
            with cb:
                if st.button("🔄 استبدال الكل | Replace All",type="secondary",use_container_width=True):
                    st.session_state["amz_confirm_replace_inv"]=True
            if st.session_state.get("amz_confirm_replace_inv"):
                st.warning("⚠️ هيمسح الكل ويرفع الجديد؟")
                cy,cn=st.columns(2)
                if cy.button("✅ نعم",key="amz_yes_rep_inv"):
                    n=do_inv_upload(replace=True)
                    st.session_state["amz_confirm_replace_inv"]=False
                    st.success(f"✅ تم الاستبدال — {n} صف"); st.rerun()
                if cn.button("❌ لا",key="amz_no_rep_inv"):
                    st.session_state["amz_confirm_replace_inv"]=False; st.rerun()
        except Exception as e:
            st.error(f"❌ {e}")

    st.divider()
    st.subheader("📋 بيانات المخزون الحالية | Current Inventory")
    if not inv_map:
        st.info("لم يُرفع ملف مخزون بعد | No inventory uploaded yet.")
    else:
        if excluded_wh: st.info(f"⚙️ مستثنى | Excluded: **{', '.join(sorted(excluded_wh))}**")
        srch=st.text_input("🔍 بحث ASIN",key="amz_srch_inv")
        raw_inv=get_cached(inventory_sheet)
        df_inv_dl=pd.DataFrame(raw_inv[1:],columns=raw_inv[0])
        c1,c2=st.columns(2)
        with c1: dl_btn(df_inv_dl,"amz_inventory")
        with c2:
            if st.button("🗑️ مسح الكل",type="secondary",key="amz_btn_clear_inv",use_container_width=True):
                st.session_state["amz_confirm_clear_inv"]=True
        confirm_clear("clear_inv",inventory_sheet,"المخزون | Inventory")
        filtered_inv={k:v for k,v in inv_map.items() if not srch or srch.strip().upper() in k}
        st.write(f"**ASINs: {len(filtered_inv)}**")
        for asin_key,info in filtered_inv.items():
            c_img,c_inf=st.columns([1,6])
            with c_img: show_img(info["img"],70)
            with c_inf:
                st.markdown(f"**ASIN:** `{info['asin']}`")
                st.caption(f"FNSKU: `{info['fnsku']}` | MSKU: `{info['msku']}`")
                total=info["total_stock"]; unsell=info["unsellable"]; sales=info["sales"]
                st.markdown(f"📦 **SELLABLE:** **{total}**"+(f" | ⚠️ **UNSELLABLE:** {unsell}" if unsell else "")+f" | 📈 **مبيع شهري:** **{sales}**")
                badges=[]
                for wh_cond,stk in sorted(info["warehouses"].items()):
                    wh,cond=wh_cond.split("|",1) if "|" in wh_cond else (wh_cond,"")
                    is_ex=wh.upper() in excluded_wh; is_un=cond.upper()=="UNSELLABLE"
                    bg="#4b1010" if is_ex or is_un else "#1e3a5f"
                    color="#fca5a5" if is_ex or is_un else "#93c5fd"
                    strike="text-decoration:line-through;" if is_ex else ""
                    badges.append(f'<span class="wh-badge" style="background:{bg};color:{color};{strike}">{wh}({cond[:4]}): {stk}</span>')
                st.markdown("🏭 "+"".join(badges),unsafe_allow_html=True)
                st.caption(f"📅 {info['date']}")
            st.divider()

# ══ TAB 10 — مراجعة المخزون ══
with tab10:
    st.subheader("🔴 مراجعة المخزون | Stock Review")
    settings_sr=load_settings()
    delay_sr=int(settings_sr.get("schedule_delay_days","3") or 3)
    cov_days_sr=int(settings_sr.get("schedule_coverage_days","15") or 15)
    today_sr=datetime.now().date()
    d1_sr,d2_sr,d3_sr=today_sr-timedelta(days=1),today_sr-timedelta(days=2),today_sr-timedelta(days=3)
    day_dates_sr=[d1_sr,d2_sr,d3_sr]
    day_labels_sr=[f"أمس ({d1_sr.strftime('%m-%d')})",f"أول أمس ({d2_sr.strftime('%m-%d')})",f"3 أيام ({d3_sr.strftime('%m-%d')})"]
    if not inv_map:
        st.info("ارفع ملف المخزون أولاً | Upload Inventory first")
    else:
        multi_counts_sr=build_daily_sales_counts(day_dates_sr)
        review_rows=[]
        for msku_up,info in inv_map.items():
            stock=info.get("total_stock",0); sales_m=info.get("sales",0)
            eff_avg=sales_m/30 if sales_m>0 else 0
            days_so=round(stock/eff_avg) if eff_avg>0 else 9999
            if days_so>=cov_days_sr and eff_avg>0: continue
            if eff_avg==0 and stock>0: continue
            day_counts=multi_counts_sr.get(msku_up,{d:0 for d in day_dates_sr})
            review_rows.append({
                "msku":info["msku"],"msku_up":msku_up,"stock":stock,"sales_month":sales_m,
                "img":info["img"],"days_to_stockout":days_so,
                "suggested_qty":max(round(eff_avg*18),1) if eff_avg>0 else 0,
                "day_counts":day_counts,
            })
        transferred=st.session_state.get("amz_transferred_skus",[])
        existing_in_review={r["msku_up"] for r in review_rows}
        for tr in transferred:
            if tr["msku_up"] not in existing_in_review:
                avg_tr=tr.get("effective_avg",0)
                review_rows.append({
                    "msku":tr["msku"],"msku_up":tr["msku_up"],"stock":tr["stock"],
                    "sales_month":tr["sales_month"],"img":tr["img"],
                    "days_to_stockout":tr.get("days_to_stockout",0),
                    "suggested_qty":round(avg_tr*18) if avg_tr>0 else 0,
                    "day_counts":tr.get("day_counts",{d:0 for d in day_dates_sr}),
                    "_transferred":True,
                })
        review_rows.sort(key=lambda r: r["days_to_stockout"])
        if not review_rows:
            st.success("✅ لا توجد MSKUs محتاجة مراجعة | No MSKUs need stock review")
        else:
            df_sr2=pd.DataFrame([{"MSKU":r["msku"],"Yesterday":r["day_counts"].get(d1_sr,0),
                "Day Before":r["day_counts"].get(d2_sr,0),"3 Days":r["day_counts"].get(d3_sr,0),
                "Stock":r["stock"],"Monthly Sales":r["sales_month"],
                "Suggested Qty":r["suggested_qty"],"Days to Stockout":r["days_to_stockout"]} for r in review_rows])
            c1,c2=st.columns(2)
            with c1: dl_btn(df_sr2,"amz_stock_review")
            with c2: st.error(f"🔴 MSKUs محتاجة مراجعة: {len(review_rows)}")
            for r in review_rows:
                c_img,c_inf=st.columns([1,6])
                with c_img: show_img(r["img"],70)
                with c_inf:
                    st.markdown(f"**MSKU:** `{r['msku']}`")
                    if r.get("_transferred"):
                        st.markdown('<span style="background:#7c3aed;color:white;border-radius:6px;padding:2px 10px;font-size:11px;">📌 مرحّل من تاب المبيعات — محتاج جدولة</span>',unsafe_allow_html=True)
                    st.markdown(f"📦 **مخزون:** {r['stock']} | 📈 **شهري:** {r['sales_month']}")
                    st.markdown("🛒 "+render_day_counts_md(r["day_counts"],day_dates_sr,day_labels_sr))
                    st.markdown(f"💡 **كمية مقترحة:** **{r['suggested_qty']}** | ⏳ **نفاد خلال:** {r['days_to_stockout']} يوم")
                    badge_text,badge_color,_=schedule_coverage_badge(r["msku"],r["days_to_stockout"],delay_sr)
                    st.markdown(f'<span style="background:{badge_color};color:white;border-radius:6px;padding:3px 10px;font-size:12px;">{badge_text}</span>',unsafe_allow_html=True)
                    render_recent_expired_note(r["msku"])
                    for note in get_unavailable_ordered_note(r["msku"]): st.caption(note)
                st.divider()
        st.divider()
        st.subheader("⛔ مخزون منتهي بالكامل | Completely Out of Stock")
        missing_rows=compute_missing_inventory_rows(day_dates_sr)
        if not missing_rows:
            st.success("✅ لا يوجد MSKUs خارجة عن المخزون")
        else:
            df_miss=pd.DataFrame([{"MSKU":r["msku"],"Yesterday":r["day_counts"].get(d1_sr,0),
                "Day Before":r["day_counts"].get(d2_sr,0),"3 Days":r["day_counts"].get(d3_sr,0),
                "Estimated Monthly Sales":r["est_monthly_sales"]} for r in missing_rows])
            c1,c2=st.columns(2)
            with c1: dl_btn(df_miss,"amz_out_of_stock",key="amz_dlbtn_oos_t10")
            with c2: st.error(f"⛔ MSKUs منتهية: {len(missing_rows)}")
            for r in missing_rows:
                c_img,c_inf=st.columns([1,6])
                with c_img: show_img(r["img"],70)
                with c_inf:
                    st.markdown(f"**MSKU:** `{r['msku']}`")
                    st.error("⛔ مخزونه انتهى — مش موجود في ملف المخزون")
                    st.markdown("🛒 "+render_day_counts_md(r["day_counts"],day_dates_sr,day_labels_sr))
                    st.markdown(f"📈 **مبيع شهري تقديري:** **{r['est_monthly_sales']}**")
                    badge_text,badge_color,_=schedule_coverage_badge(r["msku"],0,delay_sr)
                    st.markdown(f'<span style="background:{badge_color};color:white;border-radius:6px;padding:3px 10px;font-size:12px;">{badge_text}</span>',unsafe_allow_html=True)
                    render_recent_expired_note(r["msku"])
                    for note in get_unavailable_ordered_note(r["msku"]): st.caption(note)
                st.divider()

# ══ TAB 11 — منتهية ══
with tab11:
    st.subheader("🗂️ الجدولة منتهية الصلاحية | Expired Schedule")
    data_ex=get_cached(expired_sheet)
    if len(data_ex)<=1:
        st.info("لا يوجد منتهي | No expired items.")
    else:
        rows_ex=data_ex[1:]
        df_ex=pd.DataFrame(rows_ex,columns=data_ex[0])
        c1,c2=st.columns(2)
        with c1: dl_btn(df_ex,"amz_expired")
        with c2:
            if st.button("🗑️ مسح الكل",type="secondary",key="amz_btn_clear_ex",use_container_width=True):
                st.session_state["amz_confirm_clear_ex"]=True
        confirm_clear("clear_ex",expired_sheet,"المنتهية | Expired")
        st.write(f"**الإجمالي: {len(rows_ex)}**")
        for i,row in enumerate(rows_ex,start=2):
            while len(row)<10: row.append("")
            asn,msku_r,asin_r,fnsku,qty,sd,img,dadd,dexp=row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8]
            c_img,c_info,c_del=st.columns([1,5,1])
            with c_img: show_img(img,70)
            with c_info:
                st.markdown(f"**ASN:** `{asn}` | **MSKU:** `{msku_r}`")
                if asin_r: st.caption(f"ASIN: `{asin_r}`")
                show_sku_info(msku_r)
                st.markdown(f"**Qty:** {qty}")
                st.caption(f"📅 Schedule: {sd} | 🗂️ Expired: {dexp}")
            with c_del:
                if st.button("🗑️",key=f"amz_del_ex_{i}"):
                    safe_delete(expired_sheet,i); st.rerun()
            st.divider()

# ══ TAB 12 — الإعدادات ══
with tab12:
    st.subheader("⚙️ الإعدادات | Settings")
    st.caption("الإعدادات محفوظة في Google Sheets وتبقى بعد الإغلاق")
    current_settings=load_settings()
    st.markdown("### 🏭 المستودعات المستثناة | Excluded Warehouses")
    all_wh=sorted({r[3].strip() for r in get_cached(inventory_sheet)[1:] if len(r)>3 and r[3].strip()})
    current_ex_str=current_settings.get("excluded_warehouses","")
    current_ex_list=[w.strip() for w in current_ex_str.split(",") if w.strip()]
    if all_wh:
        selected_ex=st.multiselect("اختر المستودعات المستثناة:",
            options=all_wh,default=[w for w in current_ex_list if w in all_wh],key="amz_wh_multi")
    else:
        st.info("ارفع ملف المخزون أولاً")
        manual=st.text_input("أو اكتب يدوياً:",value=current_ex_str,key="amz_wh_manual")
        selected_ex=[w.strip() for w in manual.split(",") if w.strip()]
    if st.button("💾 حفظ الإعدادات | Save Settings",type="primary"):
        save_setting("excluded_warehouses",",".join(selected_ex))
        st.success("✅ تم الحفظ — ستُطبَّق عند إعادة التحميل"); st.rerun()
    st.divider()
    st.markdown("### ⏳ مدة وصول المخزون بعد الجدولة | Arrival Delay Days")
    current_delay=int(current_settings.get("schedule_delay_days","3") or 3)
    new_delay=st.number_input("عدد الأيام | Delay Days",min_value=0,max_value=30,value=current_delay,step=1,key="amz_delay_input")
    if st.button("💾 حفظ | Save Delay",key="amz_save_delay"):
        save_setting("schedule_delay_days",str(new_delay)); st.success("✅ تم الحفظ"); st.rerun()
    st.divider()
    st.markdown("### 📅 عدد أيام المبيعات المعروضة | Sales Display Days")
    current_sd=int(current_settings.get("sales_display_days","7") or 7)
    new_sd=st.number_input("عدد الأيام | Days",min_value=1,max_value=30,value=current_sd,step=1,key="amz_sd_input")
    if st.button("💾 حفظ | Save Sales Days",key="amz_save_sd"):
        save_setting("sales_display_days",str(new_sd)); st.success("✅ تم الحفظ"); st.rerun()
    st.divider()
    st.markdown("### 📦 أيام تغطية الجدولة | Schedule Coverage Days")
    current_cov=int(current_settings.get("schedule_coverage_days","15") or 15)
    new_cov=st.number_input("أيام التغطية | Coverage Days",min_value=5,max_value=90,value=current_cov,step=1,key="amz_cov_input")
    if st.button("💾 حفظ | Save Coverage",key="amz_save_cov"):
        save_setting("schedule_coverage_days",str(new_cov)); st.success("✅ تم الحفظ"); st.rerun()
    if excluded_wh:
        st.divider()
        st.warning(f"🚫 مستودعات مستثناة حالياً: **{', '.join(sorted(excluded_wh))}**")
    if inv_map and all_wh:
        st.divider()
        st.markdown("### 🏭 ملخص المستودعات | Warehouse Summary")
        wh_totals={}
        for info in inv_map.values():
            for wh_cond,stk in info["warehouses"].items():
                wh=wh_cond.split("|")[0]
                wh_totals[wh]=wh_totals.get(wh,0)+stk
        wh_df=pd.DataFrame([(wh,stk,"🚫 مستثنى" if wh.upper() in excluded_wh else "✅ محسوب")
             for wh,stk in sorted(wh_totals.items())],columns=["Warehouse","Total Stock","Status"])
        st.dataframe(wh_df,use_container_width=True,hide_index=True)

# ══ TAB 13 — مراجعة المبيعات ══
with tab13:
    st.subheader("📈 مراجعة المبيعات | Sales Review")
    st.caption("ASINs مبيعاتها أمس أعلى من المعتاد والمخزون لسه كافي | Yesterday sales spike but stock still ok")
    settings_rv=load_settings()
    delay_rv=int(settings_rv.get("schedule_delay_days","3") or 3)
    cov_rv=int(settings_rv.get("schedule_coverage_days","15") or 15)
    today_rv=datetime.now().date()
    e1,e2,e3=today_rv-timedelta(days=1),today_rv-timedelta(days=2),today_rv-timedelta(days=3)
    day_dates_rv=[e1,e2,e3]
    day_labels_rv=[f"أمس ({e1.strftime('%m-%d')})",f"أول أمس ({e2.strftime('%m-%d')})",f"3 أيام ({e3.strftime('%m-%d')})"]
    if not inv_map:
        st.info("ارفع ملف المخزون أولاً | Upload Inventory first")
    else:
        multi_counts_rv=build_daily_sales_counts(day_dates_rv)
        sales_review_rows=[]
        for msku_up,info in inv_map.items():
            stock=info.get("total_stock",0); sales_m=info.get("sales",0)
            if sales_m==0: continue
            eff_avg=sales_m/30
            day_counts_rv=multi_counts_rv.get(msku_up,{d:0 for d in day_dates_rv})
            qty_yesterday=day_counts_rv.get(e1,0)
            if qty_yesterday==0: continue
            days_so_monthly=round(stock/eff_avg) if eff_avg>0 else 9999
            days_so_today=round(stock/qty_yesterday) if qty_yesterday>0 else 9999
            sales_alert=abs(qty_yesterday)*30>sales_m
            stock_alert=days_so_monthly<cov_rv
            if not sales_alert or stock_alert: continue
            valid_days={1,2,3,4,5,6,7,8,10}
            if days_so_today not in valid_days: continue
            sales_review_rows.append({
                "msku":info["msku"],"msku_up":msku_up,"stock":stock,"sales_month":sales_m,"img":info["img"],
                "days_to_stockout":days_so_monthly,"days_to_stockout_today":days_so_today,"day_counts":day_counts_rv,
            })
        sales_review_rows.sort(key=lambda r:(-r["day_counts"].get(e1,0),-r["sales_month"]))
        if not sales_review_rows:
            st.success("✅ لا توجد ASINs محتاجة مراجعة مبيعات | No ASINs need sales review")
        else:
            df_sales_rv=pd.DataFrame([{"MSKU":r["msku"],"Yesterday":r["day_counts"].get(e1,0),
                "Day Before":r["day_counts"].get(e2,0),"3 Days":r["day_counts"].get(e3,0),
                "Stock":r["stock"],"Monthly Sales":r["sales_month"],
                "Days to Stockout Today Rate":r["days_to_stockout_today"]} for r in sales_review_rows])
            c1,c2=st.columns(2)
            with c1: dl_btn(df_sales_rv,"amz_sales_review")
            with c2: st.warning(f"📈 MSKUs محتاجة مراجعة: {len(sales_review_rows)}")
            for r in sales_review_rows:
                c_img,c_inf=st.columns([1,6])
                with c_img: show_img(r["img"],70)
                with c_inf:
                    st.markdown(f"**MSKU:** `{r['msku']}`")
                    st.markdown(f"📦 **مخزون:** {r['stock']} | 📈 **شهري:** {r['sales_month']}")
                    st.markdown("🛒 "+render_day_counts_md(r["day_counts"],day_dates_rv,day_labels_rv))
                    st.markdown(f"⚡ **نفاد خلال (بيع اليوم):** {r['days_to_stockout_today']} يوم")
                    badge_text,badge_color,_=schedule_coverage_badge(r["msku"],r["days_to_stockout"],delay_rv)
                    st.markdown(f'<span style="background:{badge_color};color:white;border-radius:6px;padding:3px 10px;font-size:12px;">{badge_text}</span>',unsafe_allow_html=True)
                    render_recent_expired_note(r["msku"])
                    for note in get_unavailable_ordered_note(r["msku"]): st.caption(note)
                st.divider()
        st.divider()
        st.subheader("⛔ مخزون منتهي بالكامل | Completely Out of Stock")
        missing_rv=compute_missing_inventory_rows(day_dates_rv)
        if not missing_rv:
            st.success("✅ لا يوجد MSKUs خارجة عن المخزون")
        else:
            df_miss_rv=pd.DataFrame([{"MSKU":r["msku"],"Yesterday":r["day_counts"].get(e1,0),
                "Day Before":r["day_counts"].get(e2,0),"3 Days":r["day_counts"].get(e3,0),
                "Estimated Monthly Sales":r["est_monthly_sales"]} for r in missing_rv])
            c1,c2=st.columns(2)
            with c1: dl_btn(df_miss_rv,"amz_oos_rv",key="amz_dlbtn_oos_rv")
            with c2: st.error(f"⛔ ASINs منتهية: {len(missing_rv)}")
            for r in missing_rv:
                c_img,c_inf=st.columns([1,6])
                with c_img: show_img(r["img"],70)
                with c_inf:
                    st.markdown(f"**MSKU:** `{r['msku']}`")
                    st.error("⛔ مخزونه انتهى — مش موجود في ملف المخزون")
                    st.markdown("🛒 "+render_day_counts_md(r["day_counts"],day_dates_rv,day_labels_rv))
                    st.markdown(f"📈 **مبيع شهري تقديري:** **{r['est_monthly_sales']}**")
                    badge_text,badge_color,_=schedule_coverage_badge(r["msku"],0,delay_rv)
                    st.markdown(f'<span style="background:{badge_color};color:white;border-radius:6px;padding:3px 10px;font-size:12px;">{badge_text}</span>',unsafe_allow_html=True)
                    render_recent_expired_note(r["msku"])
                    for note in get_unavailable_ordered_note(r["msku"]): st.caption(note)
                st.divider()

# ══ TAB 14 — المبيعات ══
with tab14:
    st.subheader("🛒 المبيعات اليومية | Daily Sales")
    st.caption("رفع تقرير Inventory Ledger من أمازون | Upload Amazon Inventory Ledger report")
    with st.expander("📤 رفع بيانات المبيعات | Upload Sales Data",expanded=False):
        st.caption("أعمدة: ASIN, FNSKU, MSKU, Event Type, Quantity, Date and Time — Event Type=Shipments وQty سالب = بيع")
        upl_sales=st.file_uploader("ملف المبيعات",type=["xlsx","xls","csv"],key="amz_sales_upload")
        if upl_sales:
            try:
                df_s=pd.read_csv(upl_sales,dtype=str).fillna("") if upl_sales.name.endswith(".csv") \
                    else pd.read_excel(upl_sales,dtype=str).fillna("")
                asin_s=fnsku_s=msku_s=title_s=event_s=fc_s=qty_s=date_s=None
                for c in df_s.columns:
                    cl=c.strip().lower()
                    if cl=="asin": asin_s=c
                    if cl in ("fnsku","fulfillment-channel-sku"): fnsku_s=c
                    if cl in ("msku","seller-sku"): msku_s=c
                    if "title" in cl: title_s=c
                    if "event type" in cl or cl=="event type": event_s=c
                    if "fulfillment center" in cl: fc_s=c
                    if cl=="quantity": qty_s=c
                    if "date and time" in cl: date_s=c
                    elif cl in ("datee","date") and not date_s: date_s=c
                if not asin_s:
                    for c in df_s.columns:
                        if "asin" in c.lower(): asin_s=c; break
                if not qty_s:
                    for c in df_s.columns:
                        if "quantity" in c.lower(): qty_s=c; break
                if not date_s:
                    for c in df_s.columns:
                        if "date" in c.lower(): date_s=c; break
                if not event_s:
                    for c in df_s.columns:
                        if "event" in c.lower(): event_s=c; break
                st.info(f"📊 {len(df_s)} صف | ASIN:`{asin_s}` Event:`{event_s}` Qty:`{qty_s}` Date:`{date_s}`")
                st.dataframe(df_s.head(10),use_container_width=True,height=150)
                if st.button("🔄 رفع واستبدال | Upload & Replace",type="primary",key="amz_btn_upload_sales"):
                    dn=now_str(); to_add=[]
                    for _,row in df_s.iterrows():
                        asin  =str(row[asin_s]).strip()  if asin_s  else ""
                        fnsku =str(row[fnsku_s]).strip() if fnsku_s else ""
                        msku  =str(row[msku_s]).strip()  if msku_s  else ""
                        title =str(row[title_s]).strip()[:200] if title_s else ""
                        event =str(row[event_s]).strip() if event_s else ""
                        fc    =str(row[fc_s]).strip()    if fc_s    else ""
                        qty   =str(row[qty_s]).strip()   if qty_s   else "0"
                        dv    =str(row[date_s]).strip()  if date_s  else ""
                        pd_   =parse_excel_date(dv)
                        ds    =pd_.strftime("%Y-%m-%d %H:%M:%S") if pd_ else dv
                        if asin and asin.lower()!="nan":
                            to_add.append([asin,fnsku,msku,title,event,fc,qty,ds,dn])
                    safe_delete_all(sales_sheet)
                    hdr=["ASIN","FNSKU","MSKU","Title","Event Type","Fulfillment Center","Quantity","Date","Date Uploaded"]
                    sales_sheet.update("A1",[hdr])
                    if to_add: safe_batch_append(sales_sheet,to_add)
                    clear_cache(sales_sheet)
                    st.success(f"✅ تم رفع {len(to_add)} صف واستبدال البيانات"); st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")

    st.divider()
    settings_t14=load_settings()
    sales_days_t14=int(settings_t14.get("sales_display_days","7") or 7)
    delay_t14=int(settings_t14.get("schedule_delay_days","3") or 3)
    cov_t14=int(settings_t14.get("schedule_coverage_days","15") or 15)
    today_t14=datetime.now().date()
    sales_dates=[today_t14-timedelta(days=i) for i in range(1,sales_days_t14+1)]
    sales_labels=[]
    for i,d in enumerate(sales_dates):
        if i==0:   sales_labels.append(f"أمس ({d.strftime('%m-%d')})")
        elif i==1: sales_labels.append(f"أول أمس ({d.strftime('%m-%d')})")
        else:      sales_labels.append(f"قبل {i+1} أيام ({d.strftime('%m-%d')})")

    if not inv_map:
        st.info("ارفع ملف المخزون أولاً | Upload Inventory first")
    else:
        multi_counts_t14=build_daily_sales_counts(sales_dates)
        totals_per_day={d:sum(c.get(d,0) for c in multi_counts_t14.values()) for d in sales_dates}
        st.markdown("#### 📊 إجمالي المبيعات اليومية | Daily Sales Totals")
        total_cols=st.columns(min(len(sales_dates),7))
        for ci,(d,lbl) in enumerate(zip(sales_dates,sales_labels)):
            if ci<len(total_cols):
                with total_cols[ci]:
                    day_total=totals_per_day.get(d,0); is_y=(ci==0)
                    bg="#14532d" if (is_y and day_total>0) else "#7f1d1d" if (is_y and day_total==0) else "#172554" if day_total>0 else "#1e293b"
                    nc="#86efac" if (is_y and day_total>0) else "#fca5a5" if is_y else "#93c5fd" if day_total>0 else "#64748b"
                    border="border:2px solid #22c55e;" if (is_y and day_total>0) else "border:2px solid #ef4444;" if is_y else ""
                    st.markdown(
                        f'<div style="background:{bg};border-radius:8px;padding:8px 10px;text-align:center;margin:2px;{border}">'
                        f'<div style="font-size:10px;color:#94a3b8;">{lbl.split("(")[0].strip()}</div>'
                        f'<div style="font-size:12px;color:#6b7280;">{d.strftime("%m-%d")}</div>'
                        f'<div style="font-size:{"24" if is_y else "18"}px;font-weight:bold;color:{nc};">{day_total}</div>'
                        '</div>',unsafe_allow_html=True)
        st.divider()

        sales_tab_rows=[]
        _new_transferred=[]
        for msku_up,info in inv_map.items():
            stock=info.get("total_stock",0); sales_m=info.get("sales",0)
            day_counts=multi_counts_t14.get(msku_up,{d:0 for d in sales_dates})
            total_recent=sum(day_counts.values())
            avg_daily=(total_recent/sales_days_t14) if sales_days_t14>0 else (sales_m/30 if sales_m>0 else 0)
            eff_avg=avg_daily if avg_daily>0 else (sales_m/30 if sales_m>0 else 0)
            days_so=round(stock/eff_avg) if eff_avg>0 else 9999
            sales_tab_rows.append({
                "msku":info["msku"],"msku_up":msku_up,"stock":stock,"sales_month":sales_m,
                "img":info["img"],"day_counts":day_counts,"effective_avg":eff_avg,"days_to_stockout":days_so,
            })
        sales_tab_rows.sort(key=lambda r:-r["day_counts"].get(sales_dates[0],0) if sales_dates else 0)

        srch_t14=st.text_input("🔍 بحث ASIN",key="amz_srch_t14")
        if srch_t14.strip():
            sales_tab_rows=[r for r in sales_tab_rows if srch_t14.strip().upper() in r["msku_up"]]

        if sales_tab_rows:
            df_t14=pd.DataFrame([
                {"MSKU":r["msku"],**{sales_labels[i]:r["day_counts"].get(d,0) for i,d in enumerate(sales_dates)},
                 "مخزون":r["stock"],"مبيع شهري":r["sales_month"]} for r in sales_tab_rows])
            c1,c2=st.columns(2)
            with c1: dl_btn(df_t14,"amz_sales_daily",key="amz_dlbtn_t14")
            with c2: st.info(f"📦 ASINs: {len(sales_tab_rows)} | 📅 {sales_days_t14} يوم")

        for r in sales_tab_rows:
            c_img,c_inf=st.columns([1,7])
            with c_img: show_img(r["img"],70)
            with c_inf:
                st.markdown(f"**MSKU:** `{r['msku']}`")
                show_sku_info(r["msku"])
                y_d=sales_dates[0] if sales_dates else None
                y_cnt=r["day_counts"].get(y_d,0) if y_d else 0
                bg_y="#14532d" if y_cnt>0 else "#7f1d1d"
                nc_y="#86efac" if y_cnt>0 else "#fca5a5"
                st.markdown(
                    f'<div style="background:{bg_y};border:2px solid {nc_y};border-radius:8px;padding:6px 12px;display:inline-block;margin:4px 0;">'
                    f'<span style="color:{nc_y};font-size:14px;font-weight:bold;">{"🟢" if y_cnt>0 else "🔴"} أمس: {y_cnt}</span></div>',
                    unsafe_allow_html=True)
                other_parts=[]
                for i,d in enumerate(sales_dates):
                    if i==0: continue
                    cnt=r["day_counts"].get(d,0)
                    lbl=sales_labels[i].split("(")[0].strip()
                    other_parts.append(f'<span style="color:{"#60a5fa" if cnt>0 else "#475569"};font-size:11px;">{lbl}: <b>{cnt}</b></span>')
                if other_parts: st.markdown(" &nbsp;|&nbsp; ".join(other_parts),unsafe_allow_html=True)
                days_so_disp=r["days_to_stockout"] if r["days_to_stockout"]<9999 else "—"
                st.markdown(f"📦 **مخزون:** {r['stock']} | 📈 **شهري:** {r['sales_month']} | 📊 **يومي:** {r['effective_avg']:.1f} | ⏳ **نفاد خلال:** {days_so_disp} يوم")
                badge_text,badge_color,sched=schedule_coverage_badge(r["msku"],r["days_to_stockout"],delay_t14)
                stock_ok=r["days_to_stockout"]>=cov_t14 if r["effective_avg"]>0 else False
                un_notes=get_unavailable_ordered_note(r["msku"])
                if stock_ok and not sched:
                    cov_text=f"✅ مخزون كافٍ ({r['days_to_stockout']} يوم) — لا يحتاج جدولة الآن"
                    cov_color="#15803d"
                elif stock_ok and sched:
                    src_lbl="تشييك" if sched.get("source")=="AMZ_Check" else "مجدول"
                    cov_text=f"✅ مخزون كافٍ ({r['days_to_stockout']} يوم) + ASN {sched['asn']} بتاريخ {sched['date']} [{src_lbl}]"
                    cov_color="#15803d"
                else:
                    cov_text=badge_text; cov_color=badge_color
                st.markdown(f'<span style="background:{cov_color};color:white;border-radius:6px;padding:3px 10px;font-size:12px;">{cov_text}</span>',unsafe_allow_html=True)
                is_needs_sched_only=(not stock_ok and "محتاج جدولة" in badge_text and not sched and not un_notes)
                if is_needs_sched_only:
                    _new_transferred.append({
                        "msku":r["msku"],"msku_up":r["msku_up"],"stock":r["stock"],
                        "sales_month":r["sales_month"],"img":r["img"],
                        "effective_avg":r["effective_avg"],"days_to_stockout":r["days_to_stockout"],
                        "day_counts":r["day_counts"],
                    })
                    st.caption("📌 مرحّل لتاب مراجعة المخزون | Transferred to Stock Review tab")
                if un_notes:
                    for note in un_notes: st.caption(note)
                render_recent_expired_note(r["msku"])
            st.divider()
        st.session_state["amz_transferred_skus"]=_new_transferred

# ══ TAB 15 — تحليل الجدولة ══
with tab15:
    st.subheader("🗓️ تحليل الجدولة المقترحة | Schedule Analysis")
    st.caption("ارفع أو الصق MSKUs وهيجيلك اقتراحات جدولة | Upload or paste MSKUs for scheduling suggestions")
    if not inv_map:
        st.info("ارفع ملف المخزون أولاً | Upload Inventory first")
    else:
        method_t15=st.radio("طريقة الإدخال:",["📂 رفع ملف | Upload","✏️ لصق | Paste"],horizontal=True,key="amz_method_t15")
        analysis_mskus=[]
        if "Upload" in method_t15:
            upl_t15=st.file_uploader("ارفع Excel أو CSV (عمود MSKU)",type=["xlsx","xls","csv"],key="amz_upl_t15")
            if upl_t15:
                try:
                    df_t15=pd.read_csv(upl_t15,dtype=str).fillna("") if upl_t15.name.endswith(".csv") \
                        else pd.read_excel(upl_t15,dtype=str).fillna("")
                    msku_col_t15=None
                    for c in df_t15.columns:
                        if "msku" in c.strip().lower() or "seller-sku" in c.strip().lower(): msku_col_t15=c; break
                    if not msku_col_t15: msku_col_t15=df_t15.columns[0]
                    analysis_mskus=[str(r[msku_col_t15]).strip() for _,r in df_t15.iterrows()
                                    if str(r[msku_col_t15]).strip() and str(r[msku_col_t15]).strip().lower()!="nan"]
                    st.success(f"✅ {len(analysis_mskus)} MSKU جاهز")
                except Exception as e:
                    st.error(f"❌ {e}")
        else:
            pasted_t15=st.text_area("الصق MSKUs (كل واحد في سطر):",height=120,key="amz_paste_t15",placeholder="B0DJ76S46P\nB0BH1F3JHV")
            if pasted_t15.strip():
                analysis_mskus=[l.strip() for l in pasted_t15.strip().splitlines() if l.strip()]
                st.success(f"✅ {len(analysis_mskus)} MSKU")

        if analysis_mskus:
            st.divider()
            today_t15=datetime.now().date()
            settings_t15=load_settings()
            delay_t15=int(settings_t15.get("schedule_delay_days","3") or 3)
            sales_days_t15=int(settings_t15.get("sales_display_days","7") or 7)
            cov_t15=int(settings_t15.get("schedule_coverage_days","15") or 15)
            recent_dates_t15=[today_t15-timedelta(days=i) for i in range(1,sales_days_t15+1)]
            multi_counts_t15=build_daily_sales_counts(recent_dates_t15)
            excel_rows_t15=[]
            st.write(f"**تحليل {len(analysis_mskus)} MSKU — أيام التغطية: {cov_t15} يوم**")

            for msku_raw in analysis_mskus:
                msku_up=msku_raw.strip().upper()
                info=inv_map.get(msku_up)
                st.markdown(f"### 📦 MSKU: `{msku_raw}`")
                if not info:
                    st.error("⛔ هذا MSKU مش موجود في المخزون — مخزونه انتهى أو لم يُرفع")
                    day_counts_miss=multi_counts_t15.get(msku_up,{})
                    total_miss=sum(day_counts_miss.values())
                    if total_miss>0:
                        avg_miss=total_miss/sales_days_t15
                        est_monthly=round(avg_miss*30)
                        suggested_urgent=max(round(avg_miss*cov_t15),1)
                        urgent_date=today_t15+timedelta(days=3)
                        st.warning(f"📈 باع {total_miss} قطعة في آخر {sales_days_t15} يوم — مبيع شهري تقديري: **{est_monthly}**")
                        st.markdown(
                            f'<div style="background:#1a0000;border:1px solid #ef4444;border-left:5px solid #ef4444;border-radius:8px;padding:10px 14px;color:white;margin:6px 0;">'
                            f'🗓️ <b>جدولة مقترحة عاجلة:</b><br>'
                            f'📅 التاريخ المقترح: <b style="color:#fca5a5;">{urgent_date.strftime("%Y-%m-%d")}</b> &nbsp;|&nbsp; '
                            f'📦 الكمية المقترحة ({cov_t15} يوم): <b style="color:#fca5a5;">{suggested_urgent}</b><br>'
                            f'<span style="color:#f87171;font-size:12px;">⚠️ مخزون منتهي — يُنصح بالجدولة فوراً</span>'
                            f'</div>',unsafe_allow_html=True)
                        excel_rows_t15.append({"MSKU":msku_raw,"المخزون":0,"مبيع شهري":est_monthly,
                            "متوسط يومي":round(avg_miss,2),"نفاد خلال":"خلص",
                            "تاريخ جدولة #1":urgent_date.strftime("%Y-%m-%d"),
                            "وصول #1":(urgent_date+timedelta(days=delay_t15)).strftime("%Y-%m-%d"),
                            "كمية #1":suggested_urgent,"ملاحظة #1":"عاجل — مخزون منتهي",
                            "تاريخ جدولة #2":"","وصول #2":"","كمية #2":"","ملاحظة #2":"",
                            "تاريخ جدولة #3":"","وصول #3":"","كمية #3":"","ملاحظة #3":""})
                    else:
                        excel_rows_t15.append({"MSKU":msku_raw,"المخزون":0,"مبيع شهري":0,
                            "متوسط يومي":0,"نفاد خلال":"خلص",
                            "تاريخ جدولة #1":"","وصول #1":"","كمية #1":"","ملاحظة #1":"مخزون منتهي ولا مبيعات",
                            "تاريخ جدولة #2":"","وصول #2":"","كمية #2":"","ملاحظة #2":"",
                            "تاريخ جدولة #3":"","وصول #3":"","كمية #3":"","ملاحظة #3":""})
                    st.divider(); continue

                stock=info.get("total_stock",0); sales_m=info.get("sales",0); img=info.get("img","")
                avg_daily=sales_m/30 if sales_m>0 else 0
                day_counts_t15=multi_counts_t15.get(msku_up,{d:0 for d in recent_dates_t15})
                recent_total=sum(day_counts_t15.values())
                avg_daily_recent=(recent_total/sales_days_t15) if sales_days_t15>0 else avg_daily
                eff_avg=avg_daily_recent if avg_daily_recent>0 else avg_daily
                days_so=round(stock/eff_avg) if eff_avg>0 else 0
                stockout_date=today_t15+timedelta(days=days_so) if days_so>0 else today_t15

                c_img_t15,c_info_t15=st.columns([1,6])
                with c_img_t15: show_img(img,65)
                with c_info_t15:
                    st.markdown(f"**MSKU:** `{msku_raw}`")
                    st.markdown(f"📦 **مخزون:** **{stock}** | 📈 **شهري:** **{sales_m}** | 📊 **يومي أخير:** **{avg_daily_recent:.1f}**")
                    if eff_avg>0:
                        st.markdown(f"⏳ **متوقع النفاد:** **{days_so} يوم** ({stockout_date.strftime('%Y-%m-%d')})")
                    else:
                        st.caption("⚠️ لا توجد مبيعات مسجلة — لا يمكن تقدير يوم النفاد")

                existing_scheds=[]
                for sk in ("AMZ_Scheduled","AMZ_Check"):
                    sd=get_cached(sheets[sk])
                    if len(sd)<=1: continue
                    for row in sd[1:]:
                        while len(row)<6: row.append("")
                        if row[1].strip().upper()==msku_up:
                            d_p=parse_excel_date(row[5])
                            existing_scheds.append({"asn":row[0],"qty":row[4],"date":row[5],"parsed":d_p,"source":sk})
                existing_scheds.sort(key=lambda s:s["parsed"] or datetime.max)
                if existing_scheds:
                    st.markdown("**📋 الجدولات الحالية:**")
                    for es in existing_scheds:
                        arr_es=(es["parsed"]+timedelta(days=delay_t15)).date() if es["parsed"] else None
                        src_l="تشييك" if es["source"]=="AMZ_Check" else "مجدول"
                        st.markdown(f'<span style="background:#1e3a5f;color:#93c5fd;border-radius:6px;padding:3px 10px;font-size:12px;margin:2px;">ASN {es["asn"]} | {es["qty"]} قطعة | {es["date"]} | {src_l}{f" | وصول: {arr_es}" if arr_es else ""}</span>',unsafe_allow_html=True)

                st.markdown("---")
                st.markdown(f"**🗓️ الجدولات المقترحة — كل جدولة تغطي {cov_t15} يوم فقط:**")
                last_covered=today_t15
                if existing_scheds:
                    for es in existing_scheds:
                        if es["parsed"]:
                            arr=(es["parsed"]+timedelta(days=delay_t15)).date()
                            if arr>last_covered: last_covered=arr
                total_incoming=sum(_to_int(es["qty"]) for es in existing_scheds)
                adj_stock=stock+total_incoming
                adj_days=round(adj_stock/eff_avg) if eff_avg>0 else 999
                adj_so=today_t15+timedelta(days=adj_days)
                if existing_scheds:
                    st.caption(f"📦 بعد الجدولات الحالية: مخزون فعلي = {adj_stock} → نفاد بعد {adj_days} يوم ({adj_so.strftime('%Y-%m-%d')})")

                BUFFER=3; suggested_list=[]; running_stock=adj_stock; running_date=last_covered
                for sg_i in range(3):
                    if eff_avg<=0: break
                    sg_qty=max(round(eff_avg*cov_t15),1)
                    days_until_out=round(running_stock/eff_avg) if eff_avg>0 else 999
                    days_to_next=max(days_until_out-(delay_t15+BUFFER),1)
                    sched_date=running_date+timedelta(days=days_to_next)
                    arr_date=sched_date+timedelta(days=delay_t15)
                    stock_at_arr=max(round(running_stock-eff_avg*(arr_date-running_date).days),0)
                    note=""
                    if sg_i==0 and existing_scheds: note="⚠️ يوجد جدولة حالية — هذا اقتراح الجدولة التالية بعدها"
                    elif sg_i==0 and days_so<=cov_t15: note="🔴 المخزون قريب على الخلاص — يُنصح بالجدولة العاجلة"
                    suggested_list.append({"num":sg_i+1,"schedule_date":sched_date,"arrival_date":arr_date,
                                           "qty":sg_qty,"note":note,"stock_at_arrival":stock_at_arr})
                    running_stock=stock_at_arr+sg_qty; running_date=arr_date

                colors_sg=["#14532d","#1e3a5f","#3b0764"]
                border_sg=["#22c55e","#3b82f6","#a855f7"]
                for sg in suggested_list:
                    st.markdown(
                        f'<div style="background:{colors_sg[sg["num"]-1]};border:1px solid {border_sg[sg["num"]-1]};border-left:5px solid {border_sg[sg["num"]-1]};border-radius:8px;padding:10px 14px;color:white;margin:6px 0;">'
                        f'🗓️ <b>الجدولة {sg["num"]}:</b><br>'
                        f'📅 تاريخ الجدولة: <b style="color:#86efac;">{sg["schedule_date"].strftime("%Y-%m-%d")}</b> → وصول: <b style="color:#93c5fd;">{sg["arrival_date"].strftime("%Y-%m-%d")}</b><br>'
                        f'📦 كمية ({cov_t15} يوم): <b style="color:#c4b5fd;">{sg["qty"]}</b> | مخزون عند الوصول: <b>{sg["stock_at_arrival"]}</b>'
                        +(f'<br><span style="color:#fcd34d;font-size:12px;">📝 {sg["note"]}</span>' if sg["note"] else "")
                        +'</div>',unsafe_allow_html=True)

                render_recent_expired_note(msku_raw)
                for note in get_unavailable_ordered_note(msku_raw): st.caption(note)
                excel_row_t15={"MSKU":msku_raw,"المخزون":stock,"مبيع شهري":sales_m,
                                "متوسط يومي":round(eff_avg,2),"نفاد خلال":days_so if eff_avg>0 else "—"}
                for sg in suggested_list:
                    n=sg["num"]
                    excel_row_t15[f"تاريخ جدولة #{n}"]=sg["schedule_date"].strftime("%Y-%m-%d")
                    excel_row_t15[f"وصول #{n}"]=sg["arrival_date"].strftime("%Y-%m-%d")
                    excel_row_t15[f"كمية #{n}"]=sg["qty"]
                    excel_row_t15[f"مخزون عند الوصول #{n}"]=sg["stock_at_arrival"]
                    excel_row_t15[f"ملاحظة #{n}"]=sg["note"]
                excel_rows_t15.append(excel_row_t15)
                st.divider()

            if excel_rows_t15:
                st.divider()
                df_excel_t15=pd.DataFrame(excel_rows_t15)
                dl_btn(df_excel_t15,"amz_schedule_analysis",label="⬇️ تحميل تحليل الجدولة Excel",key="amz_dlbtn_t15_excel")

# ══ TAB 16 — مخزون بدون بيع ══
with tab16:
    st.subheader("📦 مخزون بدون بيع | Stock With No Sales")
    if not inv_map:
        st.info("ارفع ملف المخزون أولاً | Upload Inventory first")
    else:
        today_t16=datetime.now().date()
        dates_1d=[today_t16-timedelta(days=1)]
        dates_3d=[today_t16-timedelta(days=i) for i in range(1,4)]
        dates_7d=[today_t16-timedelta(days=i) for i in range(1,8)]
        all_dates_t16=list({d for d in dates_1d+dates_3d+dates_7d})
        counts_t16=build_daily_sales_counts(all_dates_t16)

        def msku_sold_in(msku_up,dates_list):
            dc=counts_t16.get(msku_up,{})
            return sum(dc.get(d,0) for d in dates_list)>0

        no_sale_1d=[]; no_sale_3d=[]; no_sale_7d=[]
        for msku_up,info in inv_map.items():
            row_t16={"msku":info["msku"],"msku_up":msku_up,"stock":info["total_stock"],"sales_month":info["sales"],"img":info["img"]}
            if not msku_sold_in(msku_up,dates_1d): no_sale_1d.append(row_t16)
            if not msku_sold_in(msku_up,dates_3d): no_sale_3d.append(row_t16)
            if not msku_sold_in(msku_up,dates_7d): no_sale_7d.append(row_t16)
        no_sale_1d.sort(key=lambda x:-x["stock"])
        no_sale_3d.sort(key=lambda x:-x["stock"])
        no_sale_7d.sort(key=lambda x:-x["stock"])

        def render_no_sale_list(rows,period_label,dl_key):
            if not rows:
                st.success(f"✅ لا يوجد MSKUs بدون مبيعات في {period_label}")
                return
            df_ns=pd.DataFrame([{"MSKU":r["msku"],"مخزون | Stock":r["stock"],"مبيع شهري | Monthly Sales":r["sales_month"]} for r in rows])
            c1,c2=st.columns(2)
            with c1: dl_btn(df_ns,dl_key,key=f"amz_dlbtn_{dl_key}")
            with c2: st.warning(f"⚠️ {len(rows)} MSKU بدون مبيعات")
            delay_ns=int(load_settings().get("schedule_delay_days","3") or 3)
            for r in rows:
                c_img,c_inf=st.columns([1,6])
                with c_img: show_img(r["img"],60)
                with c_inf:
                    st.markdown(f"**MSKU:** `{r['msku']}`")
                    st.markdown(f"📦 **مخزون:** {r['stock']} | 📈 **شهري:** {r['sales_month']}")
                    sched_ns=get_latest_schedule_info(r["msku"])
                    if sched_ns:
                        arr_ns=(sched_ns["parsed"]+timedelta(days=delay_ns)).date() if sched_ns.get("parsed") else None
                        st.caption(f"📅 ASN {sched_ns['asn']} بتاريخ {sched_ns['date']}"+(f" — وصول: {arr_ns}" if arr_ns else ""))
                    for note in get_unavailable_ordered_note(r["msku"]): st.caption(note)
                st.divider()

        sub1,sub2,sub3=st.tabs([
            f"📅 بدون مبيع أمس ({len(no_sale_1d)})",
            f"📅 بدون مبيع آخر 3 أيام ({len(no_sale_3d)})",
            f"📅 بدون مبيع آخر أسبوع ({len(no_sale_7d)})",
        ])
        with sub1: render_no_sale_list(no_sale_1d,"أمس","amz_no_sale_1d")
        with sub2: render_no_sale_list(no_sale_3d,"آخر 3 أيام","amz_no_sale_3d")
        with sub3: render_no_sale_list(no_sale_7d,"آخر أسبوع","amz_no_sale_7d")
