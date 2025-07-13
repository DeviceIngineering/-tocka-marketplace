import requests
from datetime import datetime

API_TOKEN = "f9be4985f5e3488716c040ca52b8e04c7c0f9e0b"

HEADERS = {
    "Authorization": f"Bearer {API_TOKEN}",
    "Content-Type": "application/json"
}

ORGANIZATION_UUID = "4bf22d14-4d5e-11ee-0a80-0761000a555b"
COUNTERPARTY_UUID = "5ba713c4-a31d-11ee-0a80-063f0084f98f"
STORE_UUID = "241ed919-a631-11ee-0a80-07a9000bb947"
PROJECT_UUID = "4ec39020-4e1d-11ee-0a80-00c60006dca7"

def get_product_uuid(article):
    url = "https://api.moysklad.ru/api/remap/1.2/entity/product"
    params = {
        "filter": f"article={article}",
        "limit": 1
    }
    resp = requests.get(url, headers=HEADERS, params=params)
    resp.raise_for_status()
    data = resp.json()
    rows = data.get("rows", [])
    if not rows:
        raise Exception(f"Товар с артикулом {article} не найден")
    return rows[0]["id"]

def create_order_with_position(product_uuid, quantity=1, price=10000):
    url = "https://api.moysklad.ru/api/remap/1.2/entity/customerorder"
    now_iso = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")

    body = {
        "name": f"Заказ №{now_iso}",
        "moment": now_iso,
        "organization": {
            "meta": {
                "href": f"https://api.moysklad.ru/api/remap/1.2/entity/organization/{ORGANIZATION_UUID}",
                "type": "organization"
            }
        },
        "agent": {
            "meta": {
                "href": f"https://api.moysklad.ru/api/remap/1.2/entity/counterparty/{COUNTERPARTY_UUID}",
                "type": "counterparty"
            }
        },
        "store": {
            "meta": {
                "href": f"https://api.moysklad.ru/api/remap/1.2/entity/store/{STORE_UUID}",
                "type": "store"
            }
        },
        "project": {
            "meta": {
                "href": f"https://api.moysklad.ru/api/remap/1.2/entity/project/{PROJECT_UUID}",
                "type": "project"
            }
        },
        "currency": {
            "meta": {
                "href": "https://api.moysklad.ru/api/remap/1.2/entity/currency/643",
                "type": "currency"
            }
        },
        "positions": [
            {
                "assortment": {
                    "meta": {
                        "href": f"https://api.moysklad.ru/api/remap/1.2/entity/product/{product_uuid}",
                        "type": "product"
                    }
                },
                "quantity": quantity,
                "price": price,
                "vat": 20,
                "vatEnabled": True,
                "discount": 0,
                "reserve": 0
            }
        ],
        "description": "Создан автоматически через API"
    }

    print("Тело запроса:", body)
    resp = requests.post(url, headers=HEADERS, json=body)
    resp.raise_for_status()
    return resp.json()

if __name__ == "__main__":
    try:
        article = "N315-122"
        print(f"Поиск UUID товара по артикулу {article}...")
        product_uuid = get_product_uuid(article)
        print(f"UUID товара: {product_uuid}")

        order = create_order_with_position(product_uuid, quantity=1, price=10000)
        print(f"Заказ создан! ID: {order.get('id')}")
    except requests.exceptions.HTTPError as e:
        print(f"HTTP ошибка: {e.response.status_code} - {e.response.text}")
    except Exception as e:
        print(f"Ошибка: {e}")
