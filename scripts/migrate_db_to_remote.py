#!/usr/bin/env python
import argparse
import http.cookiejar
import json
import sqlite3
import sys
import urllib.error
import urllib.request
from collections import defaultdict


def to_bool(value):
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if value is None:
        return False
    return str(value).strip().lower() in {"1", "true", "yes", "on", "tak"}


def text(value):
    if value is None:
        return ""
    return str(value)


class ApiClient:
    def __init__(self, base_url):
        self.base_url = base_url.rstrip("/")
        self.jar = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.jar))

    def request(self, method, path, payload=None):
        data = None
        headers = {"Accept": "application/json"}
        if payload is not None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"
        req = urllib.request.Request(self.base_url + path, data=data, headers=headers, method=method)
        try:
            with self.opener.open(req, timeout=60) as response:
                body = response.read().decode("utf-8")
                return response.getcode(), (json.loads(body) if body else {})
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"{method} {path} -> HTTP {exc.code}: {body}")

    def get(self, path):
        return self.request("GET", path)[1]

    def post(self, path, payload):
        return self.request("POST", path, payload)[1]

    def put(self, path, payload):
        return self.request("PUT", path, payload)[1]


def load_local_data(db_path):
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    orders = [dict(row) for row in connection.execute("SELECT * FROM orders ORDER BY created_at ASC")]
    rows = connection.execute("SELECT * FROM order_positions ORDER BY order_id, position_number ASC").fetchall()
    positions_by_order = defaultdict(list)
    for row in rows:
        positions_by_order[row["order_id"]].append(dict(row))
    connection.close()
    return orders, positions_by_order


def order_payload(row):
    return {
        "orderNumber": text(row.get("order_number")).strip(),
        "entryDate": text(row.get("entry_date")).strip(),
        "owner": text(row.get("owner")).strip(),
        "client": text(row.get("client")).strip(),
        "color": text(row.get("color")).strip(),
        "extras": text(row.get("extras")).strip(),
        "framesCount": int(row.get("frames_count") or 0),
        "sashesCount": int(row.get("sashes_count") or 0),
        "orderStatus": text(row.get("order_status") or "Dokumentacja").strip() or "Dokumentacja",
        "manualStartDate": row.get("manual_start_date"),
        "manualPlannedDate": row.get("manual_planned_date"),
    }


def position_payload(row, include_attachments=True):
    attachment = None
    if include_attachments and row.get("attachment_data"):
        attachment = {
            "name": row.get("attachment_name"),
            "mimeType": row.get("attachment_mime"),
            "dataBase64": row.get("attachment_data"),
        }

    payload = {
        "positionNumber": text(row.get("position_number")).strip(),
        "width": int(row.get("width") or 0),
        "height": int(row.get("height") or 0),
        "technology": text(row.get("technology")).strip(),
        "line": text(row.get("line")).strip(),
        "framesCount": int(row.get("position_frames_count") or 0),
        "sashesCount": int(row.get("position_sashes_count") or 0),
        "shapeRect": to_bool(row.get("shape_rect")),
        "shapeSkos": to_bool(row.get("shape_skos")),
        "shapeLuk": to_bool(row.get("shape_luk")),
        "slemieCount": int(row.get("slemie_count") or 0),
        "slupekStalyCount": int(row.get("slupek_staly_count") or 0),
        "przymykCount": int(row.get("przymyk_count") or 0),
        "niskiProgCount": int(row.get("niski_prog_count") or 0),
        "times": {
            "machining": float(row.get("machining_time") or 0.0),
            "painting": float(row.get("painting_time") or 0.0),
            "assembly": float(row.get("assembly_time") or 0.0),
        },
        "materials": {
            "wood": {"date": row.get("material_wood_date"), "toOrder": to_bool(row.get("material_wood_to_order"))},
            "corpus": {"date": row.get("material_corpus_date"), "toOrder": to_bool(row.get("material_corpus_to_order"))},
            "glass": {"date": row.get("material_glass_date"), "toOrder": to_bool(row.get("material_glass_to_order"))},
            "hardware": {"date": row.get("material_hardware_date"), "toOrder": to_bool(row.get("material_hardware_to_order"))},
            "accessories": {
                "date": row.get("material_accessories_date"),
                "toOrder": to_bool(row.get("material_accessories_to_order")),
            },
        },
        "notes": text(row.get("notes")).strip(),
        "currentDepartmentStatus": text(row.get("current_department_status") or "Dokumentacja").strip() or "Dokumentacja",
    }
    if attachment:
        payload["attachment"] = attachment
    return payload


def feedback_payload(row):
    status = text(row.get("status") or "pending").strip() or "pending"
    if status not in {"pending", "in_progress", "done"}:
        status = "pending"
    progress = row.get("progress_percent")
    try:
        progress = float(progress if progress is not None else 0.0)
    except (TypeError, ValueError):
        progress = 0.0
    if progress < 0:
        progress = 0.0
    if progress > 100:
        progress = 100.0
    return {
        "workflowStatus": status,
        "currentDepartmentStatus": text(row.get("current_department_status") or "Dokumentacja").strip() or "Dokumentacja",
        "progressPercent": progress,
        "actor": "Migracja",
    }


def main():
    parser = argparse.ArgumentParser(description="Migracja zamowien z lokalnego planner.db do aplikacji online.")
    parser.add_argument("--url", required=True, help="Adres aplikacji, np. https://twoja-apka.up.railway.app")
    parser.add_argument("--login", required=True, help="Login admina aplikacji online")
    parser.add_argument("--password", required=True, help="Haslo admina aplikacji online")
    parser.add_argument("--db", default="planner.db", help="Sciezka do lokalnej bazy SQLite")
    parser.add_argument("--skip-attachments", action="store_true", help="Pomija zalaczniki przy migracji")
    args = parser.parse_args()

    orders, positions_by_order = load_local_data(args.db)
    print(f"Lokalna baza: {len(orders)} zamowien")

    api = ApiClient(args.url)
    login_resp = api.post("/api/auth/login", {"login": args.login, "password": args.password})
    user = (login_resp or {}).get("user", {})
    print(f"Zalogowano jako: {user.get('login', 'unknown')}")

    bootstrap = api.get("/api/bootstrap")
    remote_orders = bootstrap.get("orders", [])
    remote_by_number = {text(item.get("orderNumber")).strip(): item for item in remote_orders}
    remote_positions_by_order = {}
    for item in remote_orders:
        by_number = {}
        for position in item.get("positions", []):
            by_number[text(position.get("positionNumber")).strip()] = position
        remote_positions_by_order[item["id"]] = by_number

    created_orders = 0
    updated_orders = 0
    created_positions = 0
    updated_positions = 0
    archived_order_ids = []

    for order_row in orders:
        order_number = text(order_row.get("order_number")).strip()
        if not order_number:
            continue

        payload = order_payload(order_row)
        if order_number in remote_by_number:
            remote_order = remote_by_number[order_number]
            remote_order_id = remote_order["id"]
            api.put(f"/api/orders/{remote_order_id}", payload)
            updated_orders += 1
        else:
            created = api.post("/api/orders", payload)
            remote_order_id = created["id"]
            created_orders += 1
            remote_by_number[order_number] = {"id": remote_order_id, "orderNumber": order_number, "positions": []}
            remote_positions_by_order[remote_order_id] = {}

        if to_bool(order_row.get("archived")) and text(order_row.get("order_status")) == "Zakonczone":
            archived_order_ids.append(remote_order_id)

        existing_positions = remote_positions_by_order.setdefault(remote_order_id, {})
        for pos_row in positions_by_order.get(order_row["id"], []):
            position_number = text(pos_row.get("position_number")).strip()
            if not position_number:
                continue
            pos_payload = position_payload(pos_row, include_attachments=not args.skip_attachments)
            if position_number in existing_positions:
                remote_pos_id = existing_positions[position_number]["id"]
                api.put(f"/api/positions/{remote_pos_id}", pos_payload)
                updated_positions += 1
            else:
                created_pos = api.post(f"/api/orders/{remote_order_id}/positions", pos_payload)
                remote_pos_id = created_pos["id"]
                existing_positions[position_number] = {"id": remote_pos_id, "positionNumber": position_number}
                created_positions += 1

            api.put(f"/api/positions/{remote_pos_id}/feedback", feedback_payload(pos_row))

    if archived_order_ids:
        unique_ids = sorted(set(archived_order_ids))
        chunk_size = 100
        for i in range(0, len(unique_ids), chunk_size):
            api.post("/api/orders/archive-completed", {"orderIds": unique_ids[i : i + chunk_size]})

    print("Migracja zakonczona.")
    print(f"Zamowienia: +{created_orders} nowe, {updated_orders} zaktualizowane")
    print(f"Pozycje: +{created_positions} nowe, {updated_positions} zaktualizowane")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"BLAD: {exc}")
        sys.exit(1)
