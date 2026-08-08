import os, requests
from datetime import datetime, timedelta, timezone
import pandas as pd
KST = timezone(timedelta(hours=9))
now_kst = datetime.now(KST)
print(f"리포트 실행: {now_kst}")
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
    for mp, refresh in [(Marketplaces.JP, REFRESH_TOKEN_JP), (Marketplaces.US, REFRESH_TOKEN_NA)]:
        if not refresh: continue
        try:
            creds = dict(lwa_app_id=LWA_CLIENT_ID, lwa_client_secret=LWA_CLIENT_SECRET, refresh_token=refresh)
            api = Orders(credentials=creds, marketplace=mp)
            created_after = (datetime.utcnow() - timedelta(hours=3, minutes=30)).isoformat() + "Z"
            res = api.get_orders(CreatedAfter=created_after, FulfillmentChannels=["MFN"])
            orders = res.payload.get("Orders", []) if res.payload else []
            for o in orders:
                try:
                    o["Items"] = api.get_order_items(o["AmazonOrderId"]).payload.get("OrderItems", [])
                except: o["Items"] = []
                o["Marketplace"] = mp.name
            all_orders.extend(orders)
            print(f"{mp.name}: {len(orders)}건")
        except Exception as e:
            print(f"{mp.name} 에러: {e}")
    return all_orders
def make_report(orders):
    rows = []
    for o in orders:
        purchase = o.get("PurchaseDate", "")
        kst_str = ""
        if purchase:
            try: kst_str = pd.to_datetime(purchase).tz_convert("Asia/Seoul").strftime("%Y-%m-%d %H:%M:%S")
            except: kst_str = str(purchase)
        for item in o.get("Items", []):
            rows.append({"주문번호": o.get("AmazonOrderId",""), "주문시간(KST)": kst_str, "마켓": o.get("Marketplace",""), "SKU": item.get("SellerSKU",""), "수량": item.get("QuantityOrdered","")})
    if not rows:
        rows = [{"주문번호": "최근 3시간 주문 없음", "주문시간(KST)": now_kst.strftime("%Y-%m-%d %H:%M:%S"), "SKU": "-", "수량": 0}]
    df = pd.DataFrame(rows)
    df.to_excel(f"FBM_{now_kst.strftime('%m%d_%H%M')}.xlsx", index=False)
    return "ok.xlsx", df
def send_telegram(df):
    text = f"📦 FBM 리포트 {now_kst.strftime('%m/%d %H시')}\n총 {df['주문번호'].nunique()}건\n" + "\n".join([f"• {r['주문번호']} | {r['SKU']} x{r['수량']}" for _,r in df.head(20).iterrows()])
    requests.post(f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage", json={"chat_id": TELEGRAM_CHAT_ID, "text": text[:4000]})
orders = fetch_fbm_orders()
_, df = make_report(orders)
send_telegram(df)
