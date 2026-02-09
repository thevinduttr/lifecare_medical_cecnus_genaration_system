from src.utils.load_yaml import MAX_RETRIES
from src.utils.send_error_email import capture_screenshot

def check_retry_status(current_retry_count, portal_name):
    """
    Check if current retry count has reached maximum retries.
    
    Parameters:
    - current_retry_count: Current number of retry attempts
    - portal_name: Name of the portal being retried
    
    Returns:
    - bool: True if max retries reached, False otherwise
    """
    if current_retry_count >= MAX_RETRIES:
        print(f"Max retries ({MAX_RETRIES}) reached for {portal_name}")
        return True
    return False

async def handle_max_retry_reached(page, portal_name, quotation_id, user_email=None, saved_msg=None):
    """
    Capture a screenshot when max retry is reached. This function will only capture and return the screenshot path.

    Parameters:
    - page: Current page object for screenshot
    - portal_name: Name of the portal
    - quotation_id: The quotation ID
    - user_email: Email address of the user
    - saved_msg: Optional message (kept for compatibility)

    Returns:
    - str: Screenshot path if captured, None otherwise
    """
    try:
        if page:
            screenshot_path = await capture_screenshot(page)
            print(f"Screenshot captured for {portal_name} max retry: {screenshot_path}")
            # Do not send email from here — let caller handle sending to avoid duplicates
            return screenshot_path
    except Exception as e:
        print(f"Failed to capture screenshot for {portal_name}: {e}")
    return None
