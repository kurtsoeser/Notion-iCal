from dotenv import load_dotenv
import os
from NotionClient import NotionClient

if __name__ == "__main__":
    load_dotenv()
    notion_token = os.getenv("NOTION_TOKEN")
    database_id = os.getenv("DATABASE_ID")
    if not notion_token or not database_id:
        raise SystemExit(
            "Bitte .env im Projektroot anlegen mit NOTION_TOKEN und DATABASE_ID "
            "(siehe README.MD)."
        )
    notion_client = NotionClient(notion_token)
    notion_client.get_database(database_id)
