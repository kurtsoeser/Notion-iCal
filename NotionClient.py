import requests
from datetime import date, datetime, timezone
from icalendar import Calendar, Event, vText


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

    @staticmethod
    def _rich_text_plain(prop: dict | None) -> str:
        if not prop:
            return ""
        chunks = prop.get("rich_text") or prop.get("title") or []
        return "".join(c.get("plain_text", "") for c in chunks).strip()

    @staticmethod
    def _to_utc(dt: datetime) -> datetime:
        if dt.tzinfo is None:
            return dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)

    @staticmethod
    def _parse_iso_datetime(s: str) -> datetime | None:
        if not s:
            return None
        return datetime.fromisoformat(s.replace("Z", "+00:00"))

    @staticmethod
    def _url_from_prop(prop: dict | None) -> str:
        if not prop:
            return ""
        url = prop.get("url")
        return (url or "").strip()

    @staticmethod
    def _location_from_prop(prop: dict | None) -> str:
        if not prop:
            return ""
        prop_type = prop.get("type")
        if prop_type == "rich_text":
            return NotionClient._rich_text_plain(prop)
        if prop_type == "select":
            sel = prop.get("select")
            return (sel.get("name") or "").strip() if sel else ""
        if prop_type == "place":
            place = prop.get("place") or {}
            return (place.get("name") or "").strip()
        return NotionClient._rich_text_plain(prop)

    @staticmethod
    def _notion_deep_link(https_url: str) -> str:
        if not https_url:
            return ""
        if https_url.startswith("https://"):
            return "notion://" + https_url[len("https://") :]
        if https_url.startswith("http://"):
            return "notion://" + https_url[len("http://") :]
        return https_url

    @staticmethod
    def _build_description(
        title: str,
        beschreibung: str,
        page_url: str,
        meeting_link: str,
    ) -> str:
        parts = []
        if beschreibung:
            parts.append(beschreibung)
        if page_url:
            parts.append(f"{title}\n{page_url}")
            parts.append(
                f"in Notion öffnen\n{NotionClient._notion_deep_link(page_url)}"
            )
        if meeting_link:
            parts.append(f"Meeting-Link\n{meeting_link}")
        return "\n\n".join(parts)

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
            title = self._rich_text_plain(props.get("Titel")) or "Ohne Titel"

            sel = (props.get("Kategorie") or {}).get("select")
            kategorie = sel.get("name") if sel else ""

            beschreibung = self._rich_text_plain(props.get("Beschreibung"))
            meeting_link = self._url_from_prop(props.get("Meeting Link"))
            ort = self._location_from_prop(props.get("Ort"))

            date_obj = (props.get("Datum") or {}).get("date")
            if not date_obj or not date_obj.get("start"):
                continue

            start_v = self._parse_notion_date_value(date_obj["start"])
            end_raw = date_obj.get("end")
            end_v = (
                self._parse_notion_date_value(end_raw) if end_raw else None
            )

            page_url = item["url"]
            page_id = item["id"].replace("-", "")
            summary = f"{title} [{kategorie}]" if kategorie else title

            events.append(
                {
                    "uid": f"{page_id}@notion.so",
                    "title": summary,
                    "plain_title": title,
                    "start": start_v,
                    "end": end_v,
                    "url": page_url,
                    "description": beschreibung,
                    "meeting_link": meeting_link,
                    "location": ort or "online",
                    "last_modified": self._parse_iso_datetime(
                        item.get("last_edited_time")
                    ),
                }
            )

        self.export_ical(events)

    def export_ical(self, items):
        self.cal = Calendar()
        self.cal.add("prodid", "-//Notion-iCal//kurtrocks//DE")
        self.cal.add("version", "2.0")
        self.cal.add("calscale", "GREGORIAN")
        self.cal.add("method", "PUBLISH")
        self.cal.add("name", "Notion Termine")
        self.cal.add("x-wr-calname", "Notion Termine")
        self.cal.add("x-wr-timezone", "Europe/Vienna")

        now = datetime.now(timezone.utc)

        for item in items:
            event = Event()
            event.add("summary", item["title"])
            event.add("uid", item["uid"])
            event.add("dtstamp", item.get("last_modified") or now)
            event.add("sequence", 0)
            event.add("status", "CONFIRMED")
            event.add("transp", "OPAQUE")

            page_url = item.get("url") or ""
            if page_url:
                event.add("url", vText(page_url))

            event.add("location", item.get("location") or "online")

            description = self._build_description(
                item.get("plain_title") or item["title"],
                item.get("description") or "",
                page_url,
                item.get("meeting_link") or "",
            )
            if description:
                event.add("description", description)

            st, en = item["start"], item.get("end")
            if isinstance(st, datetime):
                event.add("dtstart", self._to_utc(st))
                if en and isinstance(en, datetime):
                    event.add("dtend", self._to_utc(en))
            elif isinstance(st, date):
                event.add("dtstart", st)
                if en and isinstance(en, date):
                    event.add("dtend", en)
                else:
                    from datetime import timedelta

                    event.add("dtend", st + timedelta(days=1))

            self.cal.add_component(event)

        with open("Notion.ics", "wb") as f:
            f.write(self.cal.to_ical())
            print(f"Completely Wrote to Disk ({len(items)} events)")
