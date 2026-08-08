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
    try:
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
                        items_res = api.get_order_items(o["AmazonOrderId"])
                        items = items_res.payload.get("OrderItems", []) if items_res.payload else []
                        o["Items"] = items
                    except: o["Items"] = []
                    o["Marketplace"] = mp.name
                all_orders.extend(orders)
                print(f"{mp.name}: {len(orders)}건")
            except Exception as e:
                print(f"{mp.name} 에러: {e}")
        return all_orders
    except Exception as e:
        print(f"fetch 에러: {e}")
        return []

def make_report(orders):
    rows = []
    for o in orders:
        purchase = o.get("PurchaseDate", "")
        # 무조건 문자열로 변환 - timezone 에러 원천 차단
        if purchase:
            try:
                dt = pd.to_datetime(purchase)
                # KST로 변환 후 문자열로
                kst_str = dt.tz_convert("Asia/Seoul").strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                kst_str = str(purchase)
        else:
            kst_str = ""
        
        items = o.get("Items", [])
        if not items:
            rows.append({"주문번호": o.get("AmazonOrderId",""), "주문시간(KST)": kst_str, "마켓": o.get("Marketplace",""), "상태": o.get("OrderStatus",""), "SKU": "-", "수량": 0, "상품명": "-"})
        else:
            for item in items:
                rows.append({
                    "주문번호": o.get("AmazonOrderId",""),
                    "주문시간(KST)": kst_str,
                    "마켓": o.get("Marketplace",""),
                    "상태": o.get("OrderStatus",""),
                    "SKU": item.get("SellerSKU",""),
                    "상품명": item.get("Title","")[:80],
                    "수량": item.get("QuantityOrdered",""),
                })
    if not rows:
        rows = [{"주문번호": "최근 3시간 FBM 주문 없음", "주문시간(KST)": now_kst.strftime("%Y-%m-%d %H:%M:%S"), "SKU": "-", "수량": 0, "마켓": "-", "상태": "-", "상품명": "-"}]
    
    df = pd.DataFrame(rows)
    filename = f"FBM_Report_{now_kst.strftime('%Y%m%d_%H%M')}_KST.xlsx"
    df.to_excel(filename, index=False)
    print(f"엑셀 생성: {filename} ({len(df)}행)")
    return filename, df

def send_telegram(df):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("텔레그램 스킵")
        return
    total_orders = df["주문번호"].nunique() if "주문번호" in df.columns else len(df)
    lines = [f"📦 FBM 리포트 {now_kst.strftime('%m/%d %H시 KST')}", f"총 주문: {total_orders}건", ""]
    for _, r in df.head(20).iterrows():
        lines.append(f"• {r.get('주문번호','')} | {r.get('SKU','')} x{r.get('수량','')} | {r.get('마켓','')}")
    if len(df) > 20:
        lines.append(f"...외 {len(df)-20}건")
    text = "\n".join(lines)[:4000]
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        resp = requests.post(url, json={"chat_id": TELEGRAM_CHAT_ID, "text": text}, timeout=15)
        print(f"텔레그램 전송: {resp.status_code}")
        print(resp.text[:500])
    except Exception as e:
        print(f"텔레그램 실패: {e}")

if __name__ == "__main__":
    orders = fetch_fbm_orders()
    xlsx, df = make_report(orders)
    send_telegram(df)
