"""Simple "Hello World" email sender using the Microsoft Graph API.

This script expects configuration values in a local `.env` file.  See
`.env.example` for the list of supported variables.
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List

import requests
from dotenv import load_dotenv


@dataclass
class GraphSettings:
    """Holds the configuration required to call Microsoft Graph."""

    access_token: str
    user_id: str
    to_recipients: List[str]
    cc_recipients: List[str]
    subject: str
    body: str
    save_to_sent_items: bool
    base_url: str = "https://graph.microsoft.com/v1.0"


def load_settings() -> GraphSettings:
    """Load settings from environment variables and validate them."""

    env_path = Path(__file__).with_name(".env")
    if env_path.exists():
        load_dotenv(env_path)
    else:
        load_dotenv()

    def require_env(name: str) -> str:
        value = os.getenv(name)
        if not value:
            raise RuntimeError(
                f"Missing required environment variable: {name}. "
                "Create a .env file (see .env.example) and fill in the value."
            )
        return value

    def parse_list(name: str) -> List[str]:
        raw = os.getenv(name, "")
        recipients = [item.strip() for item in raw.split(",") if item.strip()]
        if not recipients and name == "GRAPH_TO_RECIPIENTS":
            raise RuntimeError(
                "GRAPH_TO_RECIPIENTS must contain at least one email address."
            )
        return recipients

    access_token = require_env("GRAPH_BEARER_TOKEN")
    user_id = require_env("GRAPH_USER_ID")
    to_recipients = parse_list("GRAPH_TO_RECIPIENTS")
    cc_recipients = parse_list("GRAPH_CC_RECIPIENTS")
    subject = os.getenv("GRAPH_MESSAGE_SUBJECT", "Hello from Microsoft Graph")
    body = os.getenv(
        "GRAPH_MESSAGE_BODY",
        "Hello world! This email was sent by a Python script that talks to Microsoft Graph.",
    )
    save_to_sent_items = (
        os.getenv("GRAPH_SAVE_TO_SENT_ITEMS", "false").strip().lower() == "true"
    )
    base_url = os.getenv("GRAPH_API_BASE_URL", "https://graph.microsoft.com/v1.0")

    return GraphSettings(
        access_token=access_token,
        user_id=user_id,
        to_recipients=to_recipients,
        cc_recipients=cc_recipients,
        subject=subject,
        body=body,
        save_to_sent_items=save_to_sent_items,
        base_url=base_url.rstrip("/"),
    )


def build_recipients(addresses: Iterable[str]) -> List[dict]:
    """Convert raw email strings to the payload format expected by Graph."""

    return [
        {
            "emailAddress": {
                "address": address,
            }
        }
        for address in addresses
    ]


def send_mail(settings: GraphSettings) -> None:
    """Send the email using Microsoft Graph."""

    endpoint = f"{settings.base_url}/users/{settings.user_id}/sendMail"
    payload = {
        "message": {
            "subject": settings.subject,
            "body": {
                "contentType": "Text",
                "content": settings.body,
            },
            "toRecipients": build_recipients(settings.to_recipients),
        },
        "saveToSentItems": settings.save_to_sent_items,
    }

    if settings.cc_recipients:
        payload["message"]["ccRecipients"] = build_recipients(settings.cc_recipients)

    headers = {
        "Authorization": f"Bearer {settings.access_token}",
        "Content-Type": "application/json",
    }

    response = requests.post(endpoint, headers=headers, json=payload, timeout=30)

    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        error_details = None
        try:
            error_details = response.json()
        except ValueError:
            pass
        if error_details:
            formatted = json.dumps(error_details, indent=2)
            raise RuntimeError(
                f"Graph API call failed with status {response.status_code}:\n{formatted}"
            ) from exc
        raise RuntimeError(
            f"Graph API call failed with status {response.status_code}: {response.text}"
        ) from exc

    print("Email sent successfully via Microsoft Graph.")


if __name__ == "__main__":
    send_mail(load_settings())
