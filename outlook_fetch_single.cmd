@echo off
setlocal EnableExtensions
chcp 65001 >nul

set "PYEXE="
if exist "%LocalAppData%\Programs\Python\Python310\python.exe" set "PYEXE=%LocalAppData%\Programs\Python\Python310\python.exe"
if not defined PYEXE py -V >nul 2>nul && set "PYEXE=py"
if not defined PYEXE python -V >nul 2>nul && set "PYEXE=python"

if not defined PYEXE (
  echo Python was not found. Please install Python 3 first.
  pause
  exit /b 1
)

set "TMPPY=%TEMP%\outlook_fetch_%RANDOM%_%RANDOM%.py"
set "CMD_SOURCE=%~f0"
set "PY_TARGET=%TMPPY%"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$p=$env:CMD_SOURCE; $out=$env:PY_TARGET; $marker='###PYTHON_CODE_START###'; $src=[IO.File]::ReadAllText($p,[Text.Encoding]::UTF8); $i=$src.LastIndexOf($marker); if($i -lt 0){ throw 'Python marker not found' }; $code=$src.Substring($i + $marker.Length).TrimStart([char]13,[char]10); [IO.File]::WriteAllText($out,$code,[Text.Encoding]::UTF8)"
if errorlevel 1 (
  echo Failed to prepare embedded script.
  pause
  exit /b 1
)

"%PYEXE%" "%TMPPY%" %*
set "ERR=%ERRORLEVEL%"
del "%TMPPY%" >nul 2>nul
pause
exit /b %ERR%

###PYTHON_CODE_START###
# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
from dataclasses import dataclass
import datetime as dt
import email
from email.header import decode_header, make_header
import imaplib
import json
from pathlib import Path
import re
import ssl
import sys
import textwrap
import urllib.error
import urllib.parse
import urllib.request

TOKEN_HOST = "https://login.microsoftonline.com"
IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993


class FetchError(RuntimeError):
    pass


@dataclass
class MailItem:
    index: int
    uid: str
    sender: str
    to: str
    date: str
    subject: str
    text: str
    codes: list[str]
    raw: bytes


def parse_account_line(line: str) -> tuple[str, str, str]:
    parts = line.strip().split("----", 3)
    if len(parts) != 4:
        raise FetchError("\u683c\u5f0f\u9519\u8bef\uff0c\u5e94\u4e3a\uff1a\u90ae\u7bb1----\u5bc6\u7801/\u5360\u4f4d----client_id----refresh_token")
    email_addr, _unused_password, client_id, refresh_token = [p.strip() for p in parts]
    if not email_addr or "@" not in email_addr:
        raise FetchError("\u90ae\u7bb1\u683c\u5f0f\u770b\u8d77\u6765\u4e0d\u6b63\u786e\u3002")
    if not client_id:
        raise FetchError("client_id \u4e0d\u80fd\u4e3a\u7a7a\u3002")
    if not refresh_token:
        raise FetchError("refresh_token \u4e0d\u80fd\u4e3a\u7a7a\u3002")
    return email_addr, client_id, refresh_token


def post_form(url: str, data: dict[str, str], timeout: int = 30) -> dict:
    encoded = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=encoded,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", "replace")
        try:
            payload = json.loads(body)
            message = payload.get("error_description") or payload.get("error") or body
        except json.JSONDecodeError:
            message = body
        raise FetchError(f"\u6362\u53d6 access_token \u5931\u8d25\uff1a{message}") from exc
    except urllib.error.URLError as exc:
        raise FetchError(f"\u65e0\u6cd5\u8fde\u63a5 Microsoft token \u63a5\u53e3\uff1a{exc.reason}") from exc


def get_access_token(client_id: str, refresh_token: str, tenant: str) -> tuple[str, str | None]:
    url = f"{TOKEN_HOST}/{tenant}/oauth2/v2.0/token"
    attempts = [
        {
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": "https://outlook.office.com/IMAP.AccessAsUser.All offline_access",
        },
        {
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        },
    ]
    last_error: FetchError | None = None
    for data in attempts:
        try:
            payload = post_form(url, data)
            token = payload.get("access_token")
            if not token:
                raise FetchError("Microsoft \u8fd4\u56de\u4e2d\u6ca1\u6709 access_token\u3002")
            return token, payload.get("refresh_token")
        except FetchError as exc:
            last_error = exc
    assert last_error is not None
    raise last_error


def imap_login(email_addr: str, access_token: str) -> imaplib.IMAP4_SSL:
    context = ssl.create_default_context()
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=context)

    def auth_callback(_challenge: bytes) -> bytes:
        return f"user={email_addr}\x01auth=Bearer {access_token}\x01\x01".encode("utf-8")

    try:
        imap.authenticate("XOAUTH2", auth_callback)
    except imaplib.IMAP4.error as exc:
        try:
            imap.logout()
        except Exception:
            pass
        raise FetchError(f"IMAP OAuth2 \u767b\u5f55\u5931\u8d25\uff1a{exc}") from exc
    return imap


def decode_mime(value: str | None) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def html_to_text(value: str) -> str:
    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value)
    value = re.sub(r"(?i)<br\s*/?>", "\n", value)
    value = re.sub(r"(?i)</p\s*>", "\n", value)
    value = re.sub(r"<[^>]+>", " ", value)
    return (
        value.replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
    )


def get_text_body(msg: email.message.Message) -> str:
    plain_parts: list[str] = []
    html_parts: list[str] = []

    def decode_part(part: email.message.Message) -> str:
        payload = part.get_payload(decode=True)
        if not payload:
            return ""
        charset = part.get_content_charset() or "utf-8"
        return payload.decode(charset, "replace")

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = (part.get("Content-Disposition") or "").lower()
            if "attachment" in disposition:
                continue
            if content_type == "text/plain":
                plain_parts.append(decode_part(part))
            elif content_type == "text/html":
                html_parts.append(html_to_text(decode_part(part)))
    else:
        content = decode_part(msg)
        if msg.get_content_type() == "text/html":
            html_parts.append(html_to_text(content))
        else:
            plain_parts.append(content)

    text = "\n".join(plain_parts or html_parts)
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines()]
    return "\n".join(line for line in lines if line).strip()


def extract_codes(text: str) -> list[str]:
    codes = re.findall(r"(?<!\d)\d{4,8}(?!\d)", text)
    seen: set[str] = set()
    result: list[str] = []
    for code in codes:
        if code not in seen:
            seen.add(code)
            result.append(code)
    return result[:10]


def safe_folder_name(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.@-]+", "_", value) or "mail"


def safe_filename(value: str, max_len: int = 90) -> str:
    cleaned = safe_folder_name(value).strip("._")
    return (cleaned or "no_subject")[:max_len]


def fetch_mail_items(imap: imaplib.IMAP4_SSL, mailbox: str, limit: int, unseen_only: bool) -> list[MailItem]:
    status, _ = imap.select(mailbox, readonly=True)
    if status != "OK":
        raise FetchError(f"\u65e0\u6cd5\u6253\u5f00\u90ae\u7bb1\u6587\u4ef6\u5939\uff1a{mailbox}")

    criteria = "UNSEEN" if unseen_only else "ALL"
    status, data = imap.uid("search", None, criteria)
    if status != "OK" or not data:
        raise FetchError("\u641c\u7d22\u90ae\u4ef6\u5931\u8d25\u3002")

    uids = data[0].split()
    selected = list(reversed(uids))[: max(1, limit)]
    items: list[MailItem] = []

    for idx, uid_bytes in enumerate(selected, start=1):
        status, msg_data = imap.uid("fetch", uid_bytes, "(BODY.PEEK[])")
        if status != "OK" or not msg_data:
            continue
        raw = b""
        for part in msg_data:
            if isinstance(part, tuple):
                raw += part[1]
        if not raw:
            continue

        msg = email.message_from_bytes(raw)
        text = get_text_body(msg)
        items.append(
            MailItem(
                index=idx,
                uid=uid_bytes.decode("ascii", "replace"),
                sender=decode_mime(msg.get("From")),
                to=decode_mime(msg.get("To")),
                date=decode_mime(msg.get("Date")),
                subject=decode_mime(msg.get("Subject")),
                text=text,
                codes=extract_codes(text),
                raw=raw,
            )
        )
    return items


def make_output_dir(email_addr: str) -> Path:
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path.cwd() / "mail_output" / safe_folder_name(email_addr) / timestamp


def save_eml(item: MailItem, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"{item.index:03d}_{item.uid}_{safe_filename(item.subject)}.eml"
    path.write_bytes(item.raw)
    return path


def print_mail_list(items: list[MailItem]) -> None:
    if not items:
        print("\n\u6ca1\u6709\u5339\u914d\u90ae\u4ef6\u3002")
        return
    print("\n\u90ae\u4ef6\u5217\u8868\uff1a")
    print("-" * 100)
    for item in items:
        subject = item.subject or "(\u65e0\u4e3b\u9898)"
        sender = item.sender or "(\u672a\u77e5\u53d1\u4ef6\u4eba)"
        codes = f"  \u9a8c\u8bc1\u7801: {', '.join(item.codes[:3])}" if item.codes else ""
        print(f"{item.index:>3}. {subject[:48]:<48} | {sender[:28]:<28}{codes}")
    print("-" * 100)


def print_mail_detail(item: MailItem) -> None:
    print("\n" + "=" * 100)
    print(f"\u7f16\u53f7: {item.index}")
    print(f"UID: {item.uid}")
    print(f"\u53d1\u4ef6\u4eba: {item.sender}")
    print(f"\u6536\u4ef6\u4eba: {item.to}")
    print(f"\u65f6\u95f4: {item.date}")
    print(f"\u4e3b\u9898: {item.subject}")
    if item.codes:
        print(f"\u9a8c\u8bc1\u7801: {', '.join(item.codes)}")
    print("-" * 100)
    if item.text:
        print(textwrap.fill(item.text, width=100, replace_whitespace=False))
    else:
        print("(\u6ca1\u6709\u53ef\u663e\u793a\u7684\u6587\u672c\u6b63\u6587\uff0c\u53ef\u80fd\u662f\u7eaf\u56fe\u7247\u6216\u9644\u4ef6\u90ae\u4ef6\u3002)")
    print("=" * 100)


def choose_mail(items: list[MailItem]) -> MailItem | None:
    if not items:
        print("\u5f53\u524d\u6ca1\u6709\u90ae\u4ef6\uff0c\u8bf7\u5148\u5237\u65b0/\u62c9\u53d6\u90ae\u4ef6\u3002")
        return None
    raw = input("\u8bf7\u8f93\u5165\u90ae\u4ef6\u7f16\u53f7\uff1a").strip()
    if not raw.isdigit():
        print("\u8bf7\u8f93\u5165\u6570\u5b57\u7f16\u53f7\u3002")
        return None
    number = int(raw)
    for item in items:
        if item.index == number:
            return item
    print("\u6ca1\u6709\u8fd9\u4e2a\u7f16\u53f7\u3002")
    return None


def ask_limit(default: int) -> int:
    raw = input(f"\u62c9\u53d6\u6700\u8fd1\u591a\u5c11\u5c01\uff1f\u76f4\u63a5\u56de\u8f66\u9ed8\u8ba4 {default}\uff1a").strip()
    if not raw:
        return default
    if not raw.isdigit() or int(raw) <= 0:
        print("\u8f93\u5165\u65e0\u6548\uff0c\u7ee7\u7eed\u4f7f\u7528\u5f53\u524d\u6570\u91cf\u3002")
        return default
    return int(raw)


def run_menu(imap: imaplib.IMAP4_SSL, email_addr: str, mailbox: str, limit: int, unseen_only: bool) -> None:
    output_dir = make_output_dir(email_addr)
    items = fetch_mail_items(imap, mailbox, limit, unseen_only)
    print_mail_list(items)

    while True:
        print(
            "\n\u8bf7\u9009\u62e9\u64cd\u4f5c\uff1a\n"
            "  1. \u67e5\u770b\u90ae\u4ef6\u5217\u8868\n"
            "  2. \u8f93\u5165\u7f16\u53f7\u67e5\u770b\u90ae\u4ef6\u8be6\u60c5\n"
            "  3. \u8f93\u5165\u7f16\u53f7\u4fdd\u5b58\u4e3a .eml\n"
            "  4. \u4fdd\u5b58\u5f53\u524d\u5217\u8868\u5168\u90e8\u90ae\u4ef6\n"
            "  5. \u53ea\u663e\u793a\u9a8c\u8bc1\u7801/\u6570\u5b57\u7801\n"
            "  6. \u91cd\u65b0\u62c9\u53d6/\u5237\u65b0\u90ae\u4ef6\n"
            "  7. \u4fee\u6539\u62c9\u53d6\u6570\u91cf\u5e76\u5237\u65b0\n"
            "  0. \u9000\u51fa\n"
        )
        choice = input("\u8f93\u5165\u6570\u5b57\uff1a").strip()
        if choice == "1":
            print_mail_list(items)
        elif choice == "2":
            item = choose_mail(items)
            if item:
                print_mail_detail(item)
        elif choice == "3":
            item = choose_mail(items)
            if item:
                print(f"\u5df2\u4fdd\u5b58\uff1a{save_eml(item, output_dir)}")
        elif choice == "4":
            if not items:
                print("\u5f53\u524d\u6ca1\u6709\u90ae\u4ef6\u53ef\u4fdd\u5b58\u3002")
            else:
                for item in items:
                    save_eml(item, output_dir)
                print(f"\u5df2\u4fdd\u5b58 {len(items)} \u5c01\u90ae\u4ef6\u5230\uff1a{output_dir}")
        elif choice == "5":
            found = False
            for item in items:
                if item.codes:
                    found = True
                    subject = item.subject or "(\u65e0\u4e3b\u9898)"
                    print(f"{item.index}. {subject} -> {', '.join(item.codes)}")
            if not found:
                print("\u5f53\u524d\u5217\u8868\u91cc\u6ca1\u6709\u8bc6\u522b\u5230 4-8 \u4f4d\u6570\u5b57\u7801\u3002")
        elif choice == "6":
            items = fetch_mail_items(imap, mailbox, limit, unseen_only)
            print_mail_list(items)
        elif choice == "7":
            limit = ask_limit(limit)
            items = fetch_mail_items(imap, mailbox, limit, unseen_only)
            print_mail_list(items)
        elif choice == "0":
            print("\u5df2\u9000\u51fa\u3002")
            return
        else:
            print("\u65e0\u6548\u9009\u62e9\uff0c\u8bf7\u8f93\u5165\u83dc\u5355\u91cc\u7684\u6570\u5b57\u3002")


def print_batch(items: list[MailItem], show_body: bool, save_raw: bool, output_dir: Path) -> int:
    if not items:
        print("\u6ca1\u6709\u5339\u914d\u90ae\u4ef6\u3002")
        return 0
    for item in items:
        print(f"\n[{item.index}] UID: {item.uid}")
        print(f"    From: {item.sender}")
        print(f"      To: {item.to}")
        print(f"    Date: {item.date}")
        print(f" Subject: {item.subject}")
        if item.codes:
            print(f"   \u9a8c\u8bc1\u7801: {', '.join(item.codes)}")
        if show_body and item.text:
            preview = item.text[:1500]
            print("\u6b63\u6587\u9884\u89c8:")
            print(textwrap.indent(textwrap.fill(preview, width=100, replace_whitespace=False), "    "))
        if save_raw:
            save_eml(item, output_dir)
    return len(items)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Outlook OAuth2 refresh_token \u53d6\u4ef6\u811a\u672c")
    parser.add_argument("account_line", nargs="?", help="\u90ae\u7bb1----\u5bc6\u7801/\u5360\u4f4d----client_id----refresh_token")
    parser.add_argument("-n", "--limit", type=int, default=10, help="\u62c9\u53d6\u6700\u8fd1\u591a\u5c11\u5c01\u90ae\u4ef6")
    parser.add_argument("--mailbox", default="INBOX", help="\u90ae\u7bb1\u6587\u4ef6\u5939")
    parser.add_argument("--unseen", action="store_true", help="\u53ea\u53d6\u672a\u8bfb\u90ae\u4ef6")
    parser.add_argument("--save", action="store_true", help="\u6279\u91cf\u6a21\u5f0f\u4e0b\u4fdd\u5b58 .eml \u6587\u4ef6")
    parser.add_argument("--body", action="store_true", help="\u6279\u91cf\u6a21\u5f0f\u4e0b\u663e\u793a\u6b63\u6587\u9884\u89c8")
    parser.add_argument("--no-menu", action="store_true", help="\u4e0d\u8fdb\u5165\u6570\u5b57\u83dc\u5355\uff0c\u76f4\u63a5\u8f93\u51fa")
    parser.add_argument("--tenant", default="common", help="common/consumers/organizations")
    parser.add_argument("--show-new-refresh-token", action="store_true", help="\u5982\u679c\u8fd4\u56de\u4e86\u65b0\u7684 refresh_token\uff0c\u5219\u6253\u5370\u51fa\u6765")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    account_line = args.account_line
    if not account_line:
        account_line = input("\u8bf7\u7c98\u8d34\uff1a\u90ae\u7bb1----\u5bc6\u7801/\u5360\u4f4d----client_id----refresh_token\n> ")
    try:
        email_addr, client_id, refresh_token = parse_account_line(account_line)
        print("\u6b63\u5728\u6362\u53d6 access_token...")
        access_token, new_refresh_token = get_access_token(client_id, refresh_token, args.tenant)
        if new_refresh_token:
            if args.show_new_refresh_token:
                print(f"\u65b0\u7684 refresh_token\uff1a{new_refresh_token}")
            else:
                print("\u63d0\u793a\uff1aMicrosoft \u8fd4\u56de\u4e86\u65b0\u7684 refresh_token\uff1b\u5982\u9700\u67e5\u770b\u8bf7\u52a0 --show-new-refresh-token\u3002")

        print("\u6b63\u5728\u8fde\u63a5 Outlook IMAP...")
        imap = imap_login(email_addr, access_token)
        if args.no_menu:
            out_dir = make_output_dir(email_addr)
            items = fetch_mail_items(imap, args.mailbox, args.limit, args.unseen)
            count = print_batch(items, args.body, args.save, out_dir)
            print(f"\n\u5b8c\u6210\uff1a\u5df2\u62c9\u53d6 {count} \u5c01\u90ae\u4ef6\u3002")
            if args.save:
                print(f"\u4fdd\u5b58\u76ee\u5f55\uff1a{out_dir}")
        else:
            run_menu(imap, email_addr, args.mailbox, max(1, args.limit), args.unseen)
        imap.logout()
        return 0
    except KeyboardInterrupt:
        print("\n\u5df2\u53d6\u6d88\u3002")
        return 130
    except FetchError as exc:
        print(f"\n\u9519\u8bef\uff1a{exc}", file=sys.stderr)
        return 1
    except imaplib.IMAP4.abort as exc:
        print(f"\nIMAP \u8fde\u63a5\u5df2\u65ad\u5f00\uff1a{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
