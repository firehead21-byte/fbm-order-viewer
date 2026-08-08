import os, requests, json, base64
from datetime import datetime, timedelta, timezone
import pandas as pd

KST = timezone(timedelta(hours=9))
now_kst = datetime.now(KST)

LWA_CLIENT_ID = os.getenv("LWA_CLIENT_ID")
LWA_CLIENT_SECRET = os.getenv("LWA_CLIENT_SECRET")
REFRESH_TOKEN_JP = os.getenv("REFRESH_TOKEN_JP")
REFRESH_TOKEN_NA = os.getenv("REFRESH_TOKEN_NA")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

def fetch_fbm_orders():
    from sp_api.api import Orders
    from sp_api.base import Marketplaces
    all_orders = []
    for mp, refresh in [(Marketplaces.JP, REFRESH_TOKEN_JP), (Marketplaces.US, REFRESH_TOKEN_NA), (Marketplaces.CA, REFRESH_TOKEN_NA), (Marketplaces.MX, REFRESH_TOKEN_NA)]:
        if not refresh: continue
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
                except: o["Items"] = []
            all_orders.extend(orders)
            print(f"{mp.name}: {len(orders)}건")
        except Exception as e:
            print(f"{mp.name} 에러: {e}")
    return all_orders

def make_report(orders):
    rows = []
    for o in orders:
        purchase = o.get("PurchaseDate", "")
        if purchase:
            kst_time = pd.to_datetime(purchase).tz_convert("Asia/Seoul").tz_localize(None)
        else: kst_time = ""
        for item in o.get("Items", []):
            rows.append({
                "주문번호": o.get("AmazonOrderId",""),
                "주문시간(KST)": kst_time,
                "마켓": o.get("Marketplace",""),
                "상태": o.get("OrderStatus",""),
                "SKU": item.get("SellerSKU",""),
                "상품명": item.get("Title","")[:80],
                "수량": item.get("QuantityOrdered",""),
            })
    if not rows:
        rows = [{"주문번호": "최근 3시간 FBM 주문 없음", "주문시간(KST)": now_kst.replace(tzinfo=None), "SKU": "-", "수량": 0}]
    df = pd.DataFrame(rows)
    filename = f"FBM_Report_{now_kst.strftime('%Y%m%d_%H%M')}_KST.xlsx"
    df.to_excel(filename, index=False)
    return filename, df

def send_telegram(df, drive_link):
    total_orders = df["주문번호"].nunique()
    text = f"📦 FBM 리포트 {now_kst.strftime('%m/%d %H시 KST')}\n총 주문: {total_orders}건\n" + "\n".join([f"• {r['주문번호']} | {r['SKU']} x{r['수량']} | {r['마켓']}" for _,r in df.head(20).iterrows()])
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    requests.post(url, json={"chat_id": TELEGRAM_CHAT_ID, "text": text[:4000]})

orders = fetch_fbm_orders()
xlsx, df = make_report(orders)
send_telegram(df, None)
