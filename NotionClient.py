import requests
from datetime import date, datetime
from icalendar import Calendar, Event


class NotionClient:

    def __init__(self, _token):
        self.cal = Calendar()
        self.TOKEN = _token

    @staticmethod
    def _parse_notion_date_value(s: str):
        """Notion liefert 'YYYY-MM-DD' oder ISO-Datumzeit; Rückgabe date oder datetime."""
        s = (s or "").strip()
        if not s:
            return None
        if "T" not in s:
            return datetime.strptime(s[:10], "%Y-%m-%d").date()
        normalized = s.replace("Z", "+00:00")
        return datetime.fromisoformat(normalized)

    def get_database(self, database_id):
        url = f"https://api.notion.com/v1/databases/{database_id}/query"
        headers = {
            "Authorization": f"Bearer {self.TOKEN}",
            "Notion-Version": "2022-06-28",
            "Content-Type": "application/json",
        }
        payload = {
            "sorts": [
                {
                    "direction": "ascending",
                    "timestamp": "last_edited_time",
                }
            ]
        }

        events = []

        resp = requests.post(url, headers=headers, json=payload)
        body = resp.json()
        if resp.status_code != 200 or "results" not in body:
            raise RuntimeError(
                f"Notion API antwortete mit {resp.status_code}: {body}"
            )
        items = body["results"]
        for item in items:
            props = item["properties"]
            titel_chunks = (props.get("Titel") or {}).get("title") or []
            if titel_chunks:
                title = titel_chunks[0].get("plain_text") or "Ohne Titel"
            else:
                title = "Ohne Titel"

            sel = (props.get("Kategorie") or {}).get("select")
            kategorie = sel.get("name") if sel else ""

            date_obj = (props.get("Datum") or {}).get("date")
            if not date_obj or not date_obj.get("start"):
                continue

            start_v = self._parse_notion_date_value(date_obj["start"])
            end_raw = date_obj.get("end")
            end_v = (
                self._parse_notion_date_value(end_raw) if end_raw else None
            )

            page_url = item["url"]
            summary = f"{title} [{kategorie}]" if kategorie else title

            events.append(
                {
                    "title": summary,
                    "start": start_v,
                    "end": end_v,
                    "url": page_url,
                }
            )

        self.export_ical(events)

    def export_ical(self, items):
        for item in items:
            event = Event()
            event.add("summary", item["title"])
            if item.get("url"):
                event.add("description", item["url"])

            st, en = item["start"], item.get("end")
            if isinstance(st, datetime):
                event.add("dtstart", st)
                if en and isinstance(en, datetime):
                    event.add("dtend", en)
            elif isinstance(st, date):
                event.add("dtstart", st)
                if en and isinstance(en, date):
                    event.add("dtend", en)

            self.cal.add_component(event)

        with open("Notion.ics", "wb") as f:
            f.write(self.cal.to_ical())
            print("Completely Wrote to Disk")
