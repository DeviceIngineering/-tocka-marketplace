import requests
from datetime import datetime

# === Настройки API МойСклад ===
API_TOKEN = "f9be4985f5e3488716c040ca52b8e04c7c0f9e0b"
STORE_ID   = "241ed919-a631-11ee-0a80-07a9000bb947"
BASE_URL   = "https://api.moysklad.ru/api/remap/1.2"

HEADERS = {
    "Authorization": f"Bearer {API_TOKEN}",
    "Accept-Encoding": "gzip",
    "Content-Type": "application/json"
}

def get_product_uuid(article: str) -> tuple[str | None, str | None]:
    """По артикулу возвращает (UUID, наименование) товара, или (None, None)."""
    url = f"{BASE_URL}/entity/product"
    params = {"filter": f"article={article}", "limit": 1}
    resp = requests.get(url, headers=HEADERS, params=params)
    resp.raise_for_status()
    rows = resp.json().get("rows", [])
    if not rows:
        return None, None
    item = rows[0]
    return item["id"], item.get("name")

def get_store_slots(store_id: str) -> dict[str, str]:
    """Возвращает словарь {UUID ячейки: имя ячейки} для адресного склада."""
    url = f"{BASE_URL}/entity/store/{store_id}/slots"
    resp = requests.get(url, headers=HEADERS, params={"limit": 1000})
    resp.raise_for_status()
    data = resp.json()
    return {row["id"]: row["name"] for row in data.get("rows", [])}

def get_stock_by_slot(product_uuid: str, store_id: str) -> list[dict]:
    """Возвращает список записей отчёта Остатки по ячейкам для товара на складе."""
    url = f"{BASE_URL}/report/stock/byslot/current"
    params = [
        ("filter", f"assortmentId={product_uuid}"),
        ("filter", f"storeId={store_id}"),
        ("limit", "1000")
    ]
    resp = requests.get(url, headers=HEADERS, params=params)
    resp.raise_for_status()
    return resp.json().get("rows", [])
