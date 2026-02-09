import os
from datetime import datetime
import time
import asyncio
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import smtplib

from src.utils.logger import logger
from playwright.sync_api import Page
from playwright.async_api import Page as AsyncPage

from src.utils.load_yaml import ERROR_SCREENSHOT_DIR, DEV_EMAIL, QA_EMAIL, BA_EMAIL

# Keep win32com for fallback synchronous operations
import win32com.client as win32


async def capture_screenshot(page: AsyncPage, req_id=None):
    """Capture screenshot asynchronously and include Req_Id in filename when provided"""
    # Ensure the screenshot directory exists
    os.makedirs(ERROR_SCREENSHOT_DIR, exist_ok=True)
    
    # Filename with date and time
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if req_id:
        screenshot_path = os.path.join(ERROR_SCREENSHOT_DIR, f"error_{req_id}_{timestamp}.png")
    else:
        screenshot_path = os.path.join(ERROR_SCREENSHOT_DIR, f"error_{timestamp}.png")
    
    # Capture and save the screenshot
    await page.screenshot(path=screenshot_path)
    logger.info(f"Screenshot saved at {screenshot_path}")
    return screenshot_path


def capture_screenshot_sync(page: Page, req_id=None):
    """Synchronous version of capture_screenshot for compatibility"""
    # Ensure the screenshot directory exists
    os.makedirs(ERROR_SCREENSHOT_DIR, exist_ok=True)
    
    # Filename with date and time
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if req_id:
        screenshot_path = os.path.join(ERROR_SCREENSHOT_DIR, f"error_{req_id}_{timestamp}.png")
    else:
        screenshot_path = os.path.join(ERROR_SCREENSHOT_DIR, f"error_{timestamp}.png")
    
    # Capture and save the screenshot
    page.screenshot(path=screenshot_path)
    logger.info(f"Screenshot saved at {screenshot_path}")
    return screenshot_path


async def send_error_email(screenshot_path=None, insurance_name=None, saved_msg=None, quotation_id=None, user_email=None, attachments=None):
    """Send an error email. Screenshot and attachments are optional.

    Returns True on success, False on failure.
    """
    try:
        # Default recipient
        USER_EMAIL = "support@thealtria.com"

        # Normalize values
        insurance_name = (insurance_name or "").strip()
        saved_msg = " - ".join(part.strip() for part in (saved_msg or "").split("-") if part.strip())

        # Build subject and body
        subject_parts = [p for p in [insurance_name + " INSURANCE" if insurance_name else None, saved_msg, str(quotation_id) if quotation_id else None] if p]
        subject = " - ".join(subject_parts) if subject_parts else "Error Notification"

        body = (
            f"User Email: {user_email}\nQuotation ID: {quotation_id}\n\n{saved_msg}\n\n"
            f"An error occurred during the {insurance_name or 'portal'} process."
        )

        # Build HTML body with bold error message
        html_body = (
            f"User Email: {user_email}<br>Quotation ID: {quotation_id}<br><br><strong>{saved_msg}</strong><br><br>"
            f"An error occurred during the <strong>{insurance_name or 'portal'}</strong> process."
        )

        # Prepare attachments list (only existing files)
        attach_list = []
        if screenshot_path:
            try:
                if os.path.exists(screenshot_path):
                    attach_list.append(screenshot_path)
            except Exception:
                pass

        if attachments:
            for a in attachments:
                try:
                    if a and os.path.exists(a):
                        attach_list.append(a)
                except Exception:
                    pass

        # In async context, call the synchronous Outlook API on a thread
        await asyncio.to_thread(_send_email_with_outlook,
                                subject=subject,
                                to_email=USER_EMAIL,
                                cc_email=f"{DEV_EMAIL}; {QA_EMAIL} ; {BA_EMAIL}",
                                body=body,
                                html_body=html_body,
                                attachments=attach_list if attach_list else None)

        logger.info("Error email sent successfully")
        return True

    except Exception as e:
        logger.error("Failed to send error email", exc_info=True)
        return False


async def send_aggregated_error_email(req_id, requested_portals=None, successful_portals=None, failed_portals=None, recipient_email=None, quotation_id=None):
    """Send a single aggregated admin email per request containing all recorded portal errors and screenshots.

    Format is aligned with `send_error_email` for subject/body and attachments.
    If `req_id` is provided screenshots named with the pattern `error_<req_id>_*.png` are attached.
    """
    try:
        USER_EMAIL = "maheshi.l@algospring.com"

        # Build a concise saved message summarizing failures
        if failed_portals:
            saved_msg = f"Failed portals: {', '.join(failed_portals)}"
        else:
            saved_msg = "Error Notification"

        # If there are no failed portals and the saved message is the generic one, skip sending
        if saved_msg == "Error Notification":
            logger.info(f"No failed portals for Req_Id {req_id}; skipping aggregated error email.")
            return False

        # Construct subject: saved message and optional quotation id (no 'Aggregated Error Report' prefix)
        subject_parts = [p for p in [saved_msg, str(quotation_id) if quotation_id else None] if p]
        subject = " - ".join(subject_parts) if subject_parts else "Error Notification"

        # Build body and HTML body like send_error_email
        body = (
            f"User Email: {recipient_email}\nQuotation ID: {quotation_id}\nReq_Id: {req_id}\n\n{saved_msg}\n\n"
            f"An error occurred during the portal processes. See details below."
        )

        # Gather detailed error rows from DB if available (optional enrichment)
        error_rows = []  # list of (portal, ts, err)
        try:
            from src.services.db_config.db_connect import MySQLDatabase
            from src.services.db_config.config import DB_HOST, DB_NAME, DB_USER, DB_PASSWORD

            db = MySQLDatabase(DB_HOST, DB_NAME, DB_USER, DB_PASSWORD)
            if db.connect():
                rows = db.fetch_all("SELECT portal_name, error, error_occurred FROM medical_request_error_log WHERE req_id = %s ORDER BY error_occurred", (req_id,))
                if rows:
                    for r in rows:
                        ts = r.get('error_occurred')
                        portal = r.get('portal_name')
                        err = r.get('error')
                        error_rows.append((portal, ts, err))
                db.disconnect()
        except Exception as e:
            # Non-fatal; include a note in the details
            logger.debug(f"Failed to fetch error rows for aggregated email: {e}")

        # If caller provided failed_portals, filter to only include error rows for those portals
        if failed_portals:
            def _matches_portal(portal_name, failed_list):
                pn = (portal_name or "").lower()
                for f in failed_list:
                    if not f:
                        continue
                    f_low = f.lower()
                    if f_low in pn or pn in f_low:
                        return True
                return False

            filtered = [ (p, ts, err) for (p, ts, err) in error_rows if _matches_portal(p, failed_portals) ]
            # Keep only the latest error per portal (by timestamp string), if multiple
            latest = {}
            for p, ts, err in filtered:
                key = p
                if key not in latest or str(ts) > str(latest[key][1]):
                    latest[key] = (p, ts, err)
            error_rows = list(latest.values())

        # Build HTML table only (no plain text list)
        if error_rows:
            table_rows = []
            for p, ts, err in error_rows:
                # Escape basic HTML in values
                p_html = str(p).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                err_html = str(err).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                ts_html = str(ts).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                table_rows.append(f"<tr><td style='padding:6px;border:1px solid #ddd'>{p_html}<br><small style='color:#666'>{ts_html}</small></td><td style='padding:6px;border:1px solid #ddd'>{err_html}</td></tr>")

            table_html = (
                "<table style='border-collapse:collapse;border:1px solid #ccc;width:100%'>"
                "<thead><tr><th style='padding:8px;border:1px solid #ddd;background:#f6f6f6;text-align:left'>Portal Name</th>"
                "<th style='padding:8px;border:1px solid #ddd;background:#f6f6f6;text-align:left'>Error</th></tr></thead>"
                "<tbody>" + "".join(table_rows) + "</tbody></table>"
            )

            # Add a short pointer in the plain-text body to the table
            body = body + "\n\nSee error table below."
            html_body = body.replace("\n", "<br>") + "<br><br>" + table_html
        else:
            body = body + "\n\nNo detailed error rows found in DB for the failed portals."
            html_body = body.replace("\n", "<br>")

        # Collect screenshots: prefer error_<req_id>_* but fall back to any file containing the req_id or known screenshot names
        attachments = []
        try:
            for fname in os.listdir(ERROR_SCREENSHOT_DIR):
                lower = fname.lower()
                full = os.path.join(ERROR_SCREENSHOT_DIR, fname)
                if req_id and (f"_{req_id}_" in fname or f"_{req_id}." in fname or fname.startswith(f"error_{req_id}_")) and fname.lower().endswith('.png'):
                    attachments.append(full)
                elif req_id and str(req_id) in fname and fname.lower().endswith('.png'):
                    attachments.append(full)
                elif any(key in lower for key in ("screenshot", "login_error", "error_")) and fname.lower().endswith('.png'):
                    # include as fallback (may include unrelated images) — still useful when req_id-based filenames aren't present
                    attachments.append(full)
        except Exception:
            pass

        # De-duplicate attachments
        attachments = list(dict.fromkeys(attachments))

        # If there are extra attachments passed in, include them
        # (e.g., logs or other files passed by callers)
        # Not all callers provide this – keep optional
        # Send using the Outlook thread helper
        await asyncio.to_thread(_send_email_with_outlook,
                                subject=subject,
                                to_email=USER_EMAIL,
                                cc_email=f"{DEV_EMAIL}; {QA_EMAIL} ; {BA_EMAIL}",
                                body=body,
                                html_body=html_body,
                                attachments=attachments if attachments else None)

        logger.info(f"Aggregated error email sent for Req_Id {req_id}")
        return True

    except Exception as e:
        logger.error(f"Failed to send aggregated error email for Req_Id {req_id}: {e}")
        return False


def _send_email_with_outlook(subject, to_email, cc_email, body, html_body=None, attachments=None):
    """Helper function to send email with Outlook API (synchronous)"""
    try:
        # Initialize COM for this thread
        import pythoncom
        pythoncom.CoInitialize()
        
        # Initialize Outlook application
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 means MailItem

        # Set up email details
        mail.Subject = subject
        mail.To = to_email
        mail.CC = cc_email
        mail.Body = body
        if html_body:
            mail.HTMLBody = html_body
        
        # Add attachments if any
        if attachments:
            for attachment in attachments:
                if os.path.exists(attachment):
                    mail.Attachments.Add(attachment)
        
        # Send the email
        mail.Send()
        return True
    except Exception as e:
        logger.error(f"Failed to send email with Outlook: {str(e)}", exc_info=True)
        return False
    finally:
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()