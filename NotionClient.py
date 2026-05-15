import html
import os
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
    def _links_from_prop(prop: dict | None) -> list[tuple[str, str]]:
        """URL oder rich_text mit Hyperlinks → [(Anzeigetext, URL), …]."""
        if not prop:
            return []
        prop_type = prop.get("type")
        if prop_type == "url":
            url = (prop.get("url") or "").strip()
            if url:
                return [("Link", url)]
        links: list[tuple[str, str]] = []
        for chunk in prop.get("rich_text") or []:
            text = (chunk.get("plain_text") or "").strip()
            href = chunk.get("href")
            if not href:
                href = (chunk.get("text", {}).get("link") or {}).get("url")
            if href:
                links.append((text or href, href.strip()))
        return links

    @staticmethod
    def _html_escape(text: str) -> str:
        return html.escape(text or "", quote=True)

    @staticmethod
    def _html_link(url: str, label: str, title: str = "") -> str:
        if not url:
            return ""
        title_attr = f' title="{NotionClient._html_escape(title)}"' if title else ""
        return (
            f'<a href="{NotionClient._html_escape(url)}" target="_blank"'
            f'{title_attr}>{NotionClient._html_escape(label)}</a>'
        )

    @staticmethod
    def _html_table_row(label: str, cell_html: str) -> str:
        return (
            "<tr>"
            f'<td style="font-weight:bold;background:#f3f2f1;padding:6px 10px;'
            f'width:150px;vertical-align:top;">{NotionClient._html_escape(label)}</td>'
            f'<td style="padding:6px 10px;vertical-align:top;">{cell_html}</td>'
            "</tr>"
        )

    @staticmethod
    def _build_description_plain(
        title: str,
        beschreibung: str,
        page_url: str,
        meeting_link: str,
        veranstaltungs_links: list[tuple[str, str]],
        aufzeichnung_links: list[tuple[str, str]],
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
        for label, url in veranstaltungs_links:
            parts.append(f"Veranstaltungslink\n{label}\n{url}")
        for label, url in aufzeichnung_links:
            parts.append(f"Aufzeichnung\n{label}\n{url}")
        return "\n\n".join(parts)

    @staticmethod
    def _build_description_html(
        title: str,
        beschreibung: str,
        page_url: str,
        meeting_link: str,
        veranstaltungs_links: list[tuple[str, str]],
        aufzeichnung_links: list[tuple[str, str]],
    ) -> str:
        rows = []
        if beschreibung:
            safe = NotionClient._html_escape(beschreibung).replace("\n", "<br>")
            rows.append(NotionClient._html_table_row("Beschreibung", safe))

        if page_url:
            rows.append(
                NotionClient._html_table_row(
                    "Notion",
                    NotionClient._html_link(page_url, title, title),
                )
            )
            notion_app = NotionClient._notion_deep_link(page_url)
            rows.append(
                NotionClient._html_table_row(
                    "In Notion öffnen",
                    NotionClient._html_link(
                        notion_app, "in Notion öffnen", "In der Notion-App öffnen"
                    ),
                )
            )

        if meeting_link:
            rows.append(
                NotionClient._html_table_row(
                    "Meeting-Link",
                    NotionClient._html_link(
                        meeting_link, "Teams / Meeting beitreten", meeting_link
                    ),
                )
            )

        if veranstaltungs_links:
            links_html = "<br>".join(
                NotionClient._html_link(url, label, label)
                for label, url in veranstaltungs_links
            )
            rows.append(
                NotionClient._html_table_row("Veranstaltungslink", links_html)
            )

        if aufzeichnung_links:
            links_html = "<br>".join(
                NotionClient._html_link(url, label, label)
                for label, url in aufzeichnung_links
            )
            rows.append(NotionClient._html_table_row("Aufzeichnung", links_html))

        if not rows:
            return ""

        table = (
            '<table border="1" cellpadding="0" cellspacing="0" '
            'style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;'
            'font-size:11pt;border-color:#e1dfdd;">'
            + "".join(rows)
            + "</table>"
        )
        return (
            "<!DOCTYPE HTML><html><head><meta charset=\"UTF-8\"></head>"
            f"<body>{table}</body></html>"
        )

    def _fetch_all_database_items(self, database_id):
        """Notion liefert max. 100 Einträge pro Request – alle Seiten laden."""
        url = f"https://api.notion.com/v1/databases/{database_id}/query"
        headers = {
            "Authorization": f"Bearer {self.TOKEN}",
            "Notion-Version": "2022-06-28",
            "Content-Type": "application/json",
        }
        base_payload = {
            "page_size": 100,
            "sorts": [{"property": "Datum", "direction": "ascending"}],
        }
        items = []
        cursor = None
        while True:
            payload = dict(base_payload)
            if cursor:
                payload["start_cursor"] = cursor
            resp = requests.post(url, headers=headers, json=payload)
            body = resp.json()
            if resp.status_code != 200 or "results" not in body:
                raise RuntimeError(
                    f"Notion API antwortete mit {resp.status_code}: {body}"
                )
            items.extend(body["results"])
            if not body.get("has_more"):
                break
            cursor = body.get("next_cursor")
        return items

    @staticmethod
    def _is_future(start_v: date | datetime) -> bool:
        now = datetime.now(timezone.utc)
        if isinstance(start_v, datetime):
            return NotionClient._to_utc(start_v) >= now
        return start_v >= now.date()

    def get_database(self, database_id, only_future: bool | None = None):
        if only_future is None:
            only_future = os.getenv("ONLY_FUTURE", "").lower() in (
                "1",
                "true",
                "yes",
            )

        events = []
        skipped_no_date = 0
        skipped_past = 0

        items = self._fetch_all_database_items(database_id)
        for item in items:
            props = item["properties"]
            title = self._rich_text_plain(props.get("Titel")) or "Ohne Titel"

            sel = (props.get("Kategorie") or {}).get("select")
            kategorie = sel.get("name") if sel else ""

            beschreibung = self._rich_text_plain(props.get("Beschreibung"))
            meeting_link = self._url_from_prop(props.get("Meeting Link"))
            veranstaltungs_links = self._links_from_prop(
                props.get("Veranstaltungslink")
            )
            aufzeichnung_links = self._links_from_prop(props.get("Aufzeichnung"))
            ort = self._location_from_prop(props.get("Ort"))

            date_obj = (props.get("Datum") or {}).get("date")
            if not date_obj or not date_obj.get("start"):
                skipped_no_date += 1
                continue

            start_v = self._parse_notion_date_value(date_obj["start"])
            if only_future and not self._is_future(start_v):
                skipped_past += 1
                continue
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
                    "veranstaltungs_links": veranstaltungs_links,
                    "aufzeichnung_links": aufzeichnung_links,
                    "location": ort or "online",
                    "last_modified": self._parse_iso_datetime(
                        item.get("last_edited_time")
                    ),
                }
            )

        self.export_ical(events)
        mode = "nur zukünftige" if only_future else "alle mit Datum"
        print(
            f"Notion: {len(items)} Einträge gelesen, "
            f"{len(events)} exportiert ({mode}), "
            f"{skipped_no_date} ohne Datum übersprungen"
            + (f", {skipped_past} vergangene übersprungen" if only_future else "")
        )

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

            plain_title = item.get("plain_title") or item["title"]
            desc_args = (
                plain_title,
                item.get("description") or "",
                page_url,
                item.get("meeting_link") or "",
                item.get("veranstaltungs_links") or [],
                item.get("aufzeichnung_links") or [],
            )
            plain_desc = self._build_description_plain(*desc_args)
            html_desc = self._build_description_html(*desc_args)
            if plain_desc:
                event.add("description", plain_desc)
            if html_desc:
                alt = vText(html_desc)
                alt.params["FMTTYPE"] = "text/html"
                alt.params["CHARSET"] = "UTF-8"
                event.add("x-alt-desc", alt)

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
