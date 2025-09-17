# Send mail with Microsoft Graph

This repository contains a minimal "hello world" style Python script that sends
an email through the [Microsoft Graph API](https://learn.microsoft.com/graph/use-the-api).

## Prerequisites

* Python 3.9 or later
* A Microsoft Entra ID (Azure Active Directory) application that has the
  `Mail.Send` application permission granted and consented by an administrator.
* A tenant user mailbox that will be used as the sender.

## Setup

1. Install dependencies:

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

2. Copy the environment template and fill in the values:

   ```bash
   cp .env.example .env
   ```

   Edit `.env` and add your tenant information, sender mailbox, recipients, and
   secrets. The `GRAPH_BEARER_TOKEN` value must be refreshed at least hourly
   because Microsoft Graph access tokens expire quickly. The `.env.example`
   file includes a `curl` command that you can adapt to request a fresh token
   using the client credentials flow.

3. Run the script:

   ```bash
   python send_mail.py
   ```

   If the request succeeds, the script prints a confirmation message and the
   recipient receives the "Hello world" email. In case of failure, the script
   raises an error that includes the Graph API response for easier debugging.

## Notes

* The script reads only from the `.env` file and never writes to it.
* For production usage you should build additional error handling,
  logging, and token management (for example by requesting a new access token
  automatically before every run).
