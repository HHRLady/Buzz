"""
send_packs.py
=============
Converts each HTML host pack to PDF (using headless Chrome) and emails
it to the relevant host via Zoho SMTP.

Run by GitHub Actions after the build step. Not intended for local use.

Environment variables required (set as GitHub Secrets):
    SMTP_USER                 Zoho login, e.g. leicestershire@business-buzz.org
    SMTP_PASSWORD             Zoho App Password
    SENDER_EMAIL              From address (usually same as SMTP_USER)
    SENDER_NAME               Display name, e.g. Emma - Business Buzz L&R
    HOST_EMAIL_MARKETHARBOROUGH
    HOST_EMAIL_LEICESTER
    HOST_EMAIL_LUTTERWORTH
    HOST_EMAIL_HINCKLEY
    HOST_EMAIL_LOUGHBOROUGH
    PAGES_BASE_URL            GitHub Pages base URL, e.g.
                              https://hhrlady.github.io/Buzz

Optional (defaults shown):
    SMTP_HOST                 smtp.zoho.eu
    SMTP_PORT                 587

Usage:
    python send_packs.py
    python send_packs.py --dry-run   (prints what would be sent, no emails)
"""

import argparse
import os
import smtplib
import subprocess
import sys
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date
from pathlib import Path

# ------------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------------

SCRIPT_DIR     = Path(__file__).resolve().parent
REGION_CURATED = SCRIPT_DIR / "Buzz_Region_Curated"
PACKS_DIR      = REGION_CURATED / "host_packs"

MONTH_LABEL = date.today().strftime("%B %Y")

# Map town code -> (display name, env var for host email, HTML filename)
TOWNS = {
    "MarketHarborough": ("Market Harborough", "HOST_EMAIL_MARKETHARBOROUGH", "HostPack_MarketHarborough.html"),
    "Leicester":        ("Leicester",          "HOST_EMAIL_LEICESTER",        "HostPack_Leicester.html"),
    "Lutterworth":      ("Lutterworth",        "HOST_EMAIL_LUTTERWORTH",      "HostPack_Lutterworth.html"),
    "Hinckley":         ("Hinckley",           "HOST_EMAIL_HINCKLEY",         "HostPack_Hinckley.html"),
    "Loughborough":     ("Loughborough",       "HOST_EMAIL_LOUGHBOROUGH",     "HostPack_Loughborough.html"),
}


# ------------------------------------------------------------------
# PDF CONVERSION (headless Chrome)
# ------------------------------------------------------------------

def html_to_pdf(html_path: Path, pdf_path: Path) -> bool:
    """Convert an HTML file to PDF using headless Chrome/Chromium."""
    chrome_bins = [
        "google-chrome", "google-chrome-stable",
        "chromium", "chromium-browser",
    ]
    chrome = None
    for bin_name in chrome_bins:
        result = subprocess.run(["which", bin_name], capture_output=True, text=True)
        if result.returncode == 0:
            chrome = bin_name
            break

    if not chrome:
        print(f"[WARN] No Chrome/Chromium found. Attaching HTML instead of PDF for {html_path.name}.")
        return False

    cmd = [
        chrome,
        "--headless",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        f"--print-to-pdf={pdf_path}",
        "--print-to-pdf-no-header",
        f"file://{html_path.resolve()}",
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if result.returncode == 0 and pdf_path.exists():
            return True
        print(f"[WARN] Chrome PDF conversion failed: {result.stderr[:200]}")
        return False
    except Exception as exc:
        print(f"[WARN] Chrome PDF conversion error: {exc}")
        return False


# ------------------------------------------------------------------
# EMAIL (Zoho SMTP)
# ------------------------------------------------------------------

def _env(key: str) -> str:
    val = os.environ.get(key, "").strip()
    if not val:
        raise EnvironmentError(f"Required environment variable not set: {key}")
    return val


def send_email(
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_password: str,
    from_email: str,
    from_name: str,
    to_email: str,
    subject: str,
    html_body: str,
    text_body: str,
    attachment_path: Path | None,
    attachment_name: str,
    dry_run: bool = False,
) -> bool:

    if dry_run:
        print(f"  [DRY RUN] Would send to {to_email}: {subject}")
        if attachment_path and attachment_path.exists():
            print(f"  [DRY RUN] Attachment: {attachment_name} ({attachment_path.stat().st_size // 1024}KB)")
        return True

    try:
        # Outer container (mixed = body + attachment)
        msg = MIMEMultipart("mixed")
        msg["Subject"] = subject
        msg["From"]    = f"{from_name} <{from_email}>"
        msg["To"]      = to_email

        # Text / HTML alternative pair
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(text_body, "plain", "utf-8"))
        alt.attach(MIMEText(html_body, "html",  "utf-8"))
        msg.attach(alt)

        # Attachment
        if attachment_path and attachment_path.exists():
            with open(attachment_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f'attachment; filename="{attachment_name}"')
            msg.attach(part)

        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.ehlo()
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(from_email, to_email, msg.as_string())
        return True

    except Exception as exc:
        print(f"  [FAIL] SMTP error: {exc}")
        return False


# ------------------------------------------------------------------
# EMAIL CONTENT BUILDERS
# ------------------------------------------------------------------

def _text_body(host_name: str, town_label: str, web_link: str | None) -> str:
    lines = [
        f"Hi {host_name},",
        "",
        f"Your {town_label} host pack for {MONTH_LABEL} is attached.",
        "It covers:",
        "  1. Faces to recognise this month (your regulars)",
        "  2. Visitors worth a call before the event (lapsed re-engagement)",
        "  3. Buzz Plus conversations to have on the night",
        "  4. Sponsors",
        "",
    ]
    if web_link:
        lines += [
            "You can also view it online (useful on your phone before the event):",
            f"  {web_link}",
            "",
        ]
    lines += [
        "Any questions, just reply to this email.",
        "",
        "Thanks,",
        "Emma",
        "",
        "--",
        "Business Buzz Leicestershire & Rutland",
        "This pack is for host use only. Please do not forward or share.",
    ]
    return "\n".join(lines)


def _html_body(host_name: str, town_label: str, web_link: str | None) -> str:
    teal    = "#00A19A"
    orange  = "#F39200"
    dark    = "#111827"
    muted   = "#6B7280"
    light   = "#F9FAFB"

    link_section = ""
    if web_link:
        link_section = f"""
        <tr><td style="padding:16px 32px;background:{light};border-radius:8px;margin:0 32px">
          <p style="margin:0 0 8px;font-size:13px;color:{muted};font-family:Gill Sans,Calibri,sans-serif">
            View online (useful on your phone before the event):
          </p>
          <a href="{web_link}" style="color:{teal};font-size:13px;font-family:Gill Sans,Calibri,sans-serif">{web_link}</a>
        </td></tr>
        <tr><td style="padding:8px 0"></td></tr>"""

    return f"""<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#F3F4F6;font-family:'Century Gothic','Gill Sans',Calibri,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td align="center" style="padding:32px 16px">
<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px">

  <!-- Header -->
  <tr><td style="background:{teal};border-radius:12px 12px 0 0;padding:28px 32px 22px">
    <p style="margin:0 0 6px;font-size:11px;letter-spacing:0.1em;text-transform:uppercase;color:rgba(255,255,255,0.6)">
      Business Buzz &middot; Leicestershire &amp; Rutland
    </p>
    <p style="margin:0 0 4px;font-size:22px;font-weight:700;color:#fff">{town_label} host pack</p>
    <p style="margin:0;font-size:13px;color:rgba(255,255,255,0.7)">{MONTH_LABEL}</p>
  </td></tr>

  <!-- Colour bar -->
  <tr>
    <td style="padding:0">
      <table width="100%" cellpadding="0" cellspacing="0"><tr>
        <td style="background:{teal};height:4px"></td>
        <td style="background:{orange};height:4px"></td>
        <td style="background:#D60B52;height:4px"></td>
        <td style="background:#B6BD00;height:4px"></td>
      </tr></table>
    </td>
  </tr>

  <!-- Body -->
  <tr><td style="background:#ffffff;border:1px solid #E5E7EB;border-top:none;border-radius:0 0 12px 12px;padding:28px 32px">
    <table width="100%" cellpadding="0" cellspacing="0">

      <tr><td style="padding:0 0 16px">
        <p style="margin:0;font-size:15px;color:{dark}">Hi {host_name},</p>
      </td></tr>

      <tr><td style="padding:0 0 16px">
        <p style="margin:0;font-size:14px;color:{dark};line-height:1.6">
          Your <strong>{town_label}</strong> host pack for <strong>{MONTH_LABEL}</strong> is attached.
          It covers your regulars to recognise, lapsed visitors to re-engage, Buzz Plus prospects,
          and your sponsor picture.
        </p>
      </td></tr>

      {link_section}

      <tr><td style="padding:16px 0 0">
        <p style="margin:0;font-size:14px;color:{dark};line-height:1.6">
          Any questions, just reply to this email.
        </p>
      </td></tr>

      <tr><td style="padding:24px 0 0;border-top:1px solid #E5E7EB;margin-top:24px">
        <p style="margin:0 0 4px;font-size:14px;color:{dark}">Thanks,<br><strong>Emma</strong></p>
        <p style="margin:8px 0 0;font-size:12px;color:{muted}">Business Buzz Leicestershire &amp; Rutland</p>
        <p style="margin:4px 0 0;font-size:11px;color:{muted}">
          This pack is for host use only. Please do not forward or share.
        </p>
      </td></tr>

    </table>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>"""


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(description="Send Business Buzz host packs via Zoho SMTP.")
    ap.add_argument("--dry-run", action="store_true", help="Print what would be sent without sending")
    ap.add_argument("--town", default="ALL", help="Town code or ALL")
    args = ap.parse_args()

    # Load credentials
    try:
        smtp_host     = os.environ.get("SMTP_HOST", "smtp.zoho.eu").strip()
        smtp_port     = int(os.environ.get("SMTP_PORT", "587"))
        smtp_user     = _env("SMTP_USER")
        smtp_password = _env("SMTP_PASSWORD")
        from_email    = _env("SENDER_EMAIL")
        from_name     = _env("SENDER_NAME")
        pages_base    = os.environ.get("PAGES_BASE_URL", "").strip().rstrip("/")
    except EnvironmentError as exc:
        print(f"[FAIL] {exc}")
        sys.exit(1)

    towns_to_send = {k: v for k, v in TOWNS.items()
                     if args.town.upper() == "ALL" or k == args.town}

    if not towns_to_send:
        print(f"[FAIL] Unknown town: {args.town}")
        sys.exit(1)

    ok_count = fail_count = 0

    for town_code, (town_label, email_env, html_filename) in towns_to_send.items():

        print(f"\n[TOWN] {town_label}")

        # 1. Check HTML exists
        html_path = PACKS_DIR / html_filename
        if not html_path.exists():
            print(f"  [SKIP] HTML not found: {html_path}")
            fail_count += 1
            continue

        # 2. Get host email
        try:
            to_email = _env(email_env)
        except EnvironmentError as exc:
            print(f"  [SKIP] {exc}")
            fail_count += 1
            continue

        # 3. Extract host first name from HTML (cheap grep for the greeting)
        host_name = "Host"
        try:
            content = html_path.read_text(encoding="utf-8")
            import re
            m = re.search(r"Hi\s+(\w+)\s+&#8212;", content)
            if m:
                host_name = m.group(1)
        except Exception:
            pass

        # 4. Convert HTML to PDF
        pdf_path = PACKS_DIR / html_filename.replace(".html", ".pdf")
        has_pdf  = html_to_pdf(html_path, pdf_path)
        attachment_path = pdf_path if has_pdf else html_path
        attachment_name = (html_filename.replace(".html", ".pdf") if has_pdf
                           else html_filename)

        # 5. Build web link
        web_link = None
        if pages_base:
            web_link = f"{pages_base}/{html_filename}"

        # 6. Send
        subject = f"Your {town_label} host pack \u2014 {MONTH_LABEL}"
        text_b  = _text_body(host_name, town_label, web_link)
        html_b  = _html_body(host_name, town_label, web_link)

        success = send_email(
            smtp_host      = smtp_host,
            smtp_port      = smtp_port,
            smtp_user      = smtp_user,
            smtp_password  = smtp_password,
            from_email     = from_email,
            from_name      = from_name,
            to_email       = to_email,
            subject        = subject,
            html_body      = html_b,
            text_body      = text_b,
            attachment_path= attachment_path,
            attachment_name= attachment_name,
            dry_run        = args.dry_run,
        )

        if success:
            print(f"  [OK] Sent to {to_email}")
            ok_count += 1
        else:
            fail_count += 1

    print(f"\n[DONE] {ok_count} sent, {fail_count} failed.")
    if fail_count:
        sys.exit(1)


if __name__ == "__main__":
    main()
