"""
FBM Order Viewer - 한국시간 09/12/15시 자동 리포트 (Timezone 버그 수정)
"""
import os, requests, json, base64
from datetime import datetime, timedelta, timezone
import pandas as pd

KST = timezone(timedelta(hours=9))
now_kst = datetime.now(KST)
print(f"리포트 실행: {now_kst} KST")

LWA_CLIENT_ID = os.getenv("LWA_CLIENT_ID")
LWA_CLIENT_SECRET = os.getenv("LWA_CLIENT_SECRET")
REFRESH_TOKEN_JP = os.getenv("REFRESH_TOKEN_JP")
REFRESH_TOKEN_NA = os.getenv("REFRESH_TOKEN_NA")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

def fetch_fbm_orders():
    try:
        from sp_api.api import Orders
        from sp_api.base import Marketplaces
        all_orders = []
        for mp, refresh in [(Marketplaces.JP, REFRESH_TOKEN_JP), (Marketplaces.US, REFRESH_TOKEN_NA), (Marketplaces.CA, REFRESH_TOKEN_NA), (Marketplaces.MX, REFRESH_TOKEN_NA)]:
            if not refresh:
                continue
            try:
                creds = dict(lwa_app_id=LWA_CLIENT_ID, lwa_client_secret=LWA_CLIENT_SECRET, refresh_token=refresh)
                api = Orders(credentials=creds, marketplace=mp)
                created_after = (datetime.utcnow() - timedelta(hours=3, minutes=30)).isoformat() + "Z"
                res = api.get_orders(CreatedAfter=created_after, FulfillmentChannels=["MFN"])
                orders = res.payload.get("Orders", []) if res.payload else []
                for o in orders:
                    try:
                        items = api.get_order_items(o["AmazonOrderId"]).payload.get("OrderItems", [])
                        o["Items"] = items
                        o["Marketplace"] = mp.name
                    except:
                        o["Items"] = []
                all_orders.extend(orders)
                print(f"{mp.name}: {len(orders)}건")
            except Exception as e:
                print(f"{mp.name} 에러: {e}")
        return all_orders
    except ImportError:
        print("sp_api 없음")
        return []

def make_report(orders):
    rows = []
    for o in orders:
        order_id = o.get("AmazonOrderId", "")
        purchase = o.get("PurchaseDate", "")
        status = o.get("OrderStatus", "")
        mp = o.get("Marketplace", "")
        total = o.get("OrderTotal", {}).get("Amount", "")
        # timezone 제거!
        if purchase:
            try:
                kst_time = pd.to_datetime(purchase).tz_convert(KST).tz_localize(None)
            except:
                try:
                    kst_time = pd.to_datetime(purchase).tz_localize('UTC').tz_convert(KST).tz_localize(None)
                except:
                    kst_time = purchase
        else:
            kst_time = ""
        
        for item in o.get("Items", []):
            rows.append({
                "주문번호": order_id,
                "주문시간(KST)": kst_time,
                "마켓": mp,
                "상태": status,
                "SKU": item.get("SellerSKU", ""),
                "상품명": item.get("Title", "")[:80],
                "수량": item.get("QuantityOrdered", ""),
                "가격": total,
            })
    if not rows:
        rows = [{"주문번호": "최근 3시간 FBM 주문 없음", "주문시간(KST)": now_kst.replace(tzinfo=None), "SKU": "-", "수량": 0, "마켓": "-", "상태": "-", "상품명": "-", "가격": "-"}]
    
    df = pd.DataFrame(rows)
    filename = f"FBM_Report_{now_kst.strftime('%Y%m%d_%H%M')}_KST.xlsx"
    df.to_excel(filename, index=False)
    print(f"엑셀 생성: {filename} ({len(df)}행)")
    return filename, df

def upload_to_gdrive(filepath):
    folder_id = os.getenv("GDRIVE_FOLDER_ID", "")
    sa_json_b64 = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON_B64", "")
    if not sa_json_b64:
        print("GDRIVE 스킵")
        return None
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        sa_json = json.loads(base64.b64decode(sa_json_b64).decode())
        creds = service_account.Credentials.from_service_account_info(sa_json, scopes=["https://www.googleapis.com/auth/drive"])
        service = build("drive", "v3", credentials=creds)
        file_metadata = {"name": os.path.basename(filepath)}
        if folder_id:
            file_metadata["parents"] = [folder_id]
        media = MediaFileUpload(filepath, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields="id, webViewLink").execute()
        print(f"Drive 업로드: {file.get('webViewLink')}")
        return file.get("webViewLink")
    except Exception as e:
        print(f"Drive 실패: {e}")
        return None

def send_telegram(df, drive_link):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("텔레그램 스킵")
        return
    total_orders = df["주문번호"].nunique() if "주문번호" in df.columns else 0
    summary_lines = [f"📦 FBM 리포트 {now_kst.strftime('%m/%d %H시 KST')}", f"총 주문: {total_orders}건", f"총 라인: {len(df)}개", ""]
    for _, row in df.head(20).iterrows():
        summary_lines.append(f"• {row.get('주문번호','')} | {row.get('SKU','')} x{row.get('수량','')} | {row.get('마켓','')}")
    if len(df) > 20:
        summary_lines.append(f"...외 {len(df)-20}건")
    if drive_link:
        summary_lines.append(f"\n📁 Drive: {drive_link}")
    text = "\n".join(summary_lines)[:4000]
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        r = requests.post(url, json={"chat_id": TELEGRAM_CHAT_ID, "text": text}, timeout=15)
        print(f"텔레그램: {r.status_code} {r.text[:200]}")
    except Exception as e:
        print(f"텔레그램 실패: {e}")

if __name__ == "__main__":
    orders = fetch_fbm_orders()
    xlsx, df = make_report(orders)
    link = upload_to_gdrive(xlsx)
    send_telegram(df, link)
