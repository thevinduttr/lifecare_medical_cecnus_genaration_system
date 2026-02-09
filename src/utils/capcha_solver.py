import httpx
import random
import asyncio
from .load_yaml import TWO_CAPTCHA_API_KEY

async def solve_captcha(captcha_sitekey, page_url, is_invisible=False, user_agent=None, cookies=None):
    """
    Solve reCAPTCHA v2 using 2Captcha service with the new API format.
    
    Args:
        captcha_sitekey (str): The Google reCAPTCHA sitekey
        page_url (str): The URL of the page with the CAPTCHA
        is_invisible (bool): Whether the reCAPTCHA is invisible
        user_agent (str, optional): Browser User-Agent to use
        cookies (str, optional): Cookies in format "key1=val1; key2=val2"
        
    Returns:
        str: The solved CAPTCHA token
    """
    # API endpoints from documentation
    create_task_url = 'https://api.2captcha.com/createTask'
    get_result_url = 'https://api.2captcha.com/getTaskResult'

    # Prepare the task data for RecaptchaV2TaskProxyless
    task_data = {
        "clientKey": TWO_CAPTCHA_API_KEY,
        "task": {
            "type": "RecaptchaV2TaskProxyless",
            "websiteURL": page_url,
            "websiteKey": captcha_sitekey,
            "isInvisible": is_invisible
        }
    }
    
    # Add optional parameters if provided
    if user_agent:
        task_data["task"]["userAgent"] = user_agent
    if cookies:
        task_data["task"]["cookies"] = cookies

    async with httpx.AsyncClient() as client:
        # Step 1: Create task
        response = await client.post(create_task_url, json=task_data)
        result = response.json()
        
        if result.get('errorId') != 0:
            raise Exception(f"Failed to create CAPTCHA task: {result.get('errorDescription')}")
        
        task_id = result.get('taskId')
        print(f"CAPTCHA task created successfully. Task ID: {task_id}")

        # Step 2: Poll for the task result
        max_attempts = 20  # Maximum polling attempts
        attempt = 0
        
        while attempt < max_attempts:
            await asyncio.sleep(10)  # Wait 5 seconds between polling attempts
            attempt += 1
            
            # Prepare result request
            result_data = {
                "clientKey": TWO_CAPTCHA_API_KEY,
                "taskId": task_id
            }
            
            response = await client.post(get_result_url, json=result_data)
            result = response.json()
            
            # Check the result status
            if result.get('errorId') != 0:
                raise Exception(f"Error getting task result: {result.get('errorDescription')}")
            
            if result.get('status') == 'ready':
                # CAPTCHA solved successfully
                solution = result.get('solution', {}).get('token') or result.get('solution', {}).get('gRecaptchaResponse')
                print(f"CAPTCHA solved successfully after {attempt} attempts.")

                sleep_time =  random.uniform(1, 5)
                await asyncio.sleep(sleep_time)

                return solution
            
            elif result.get('status') == 'processing':
                # CAPTCHA is still being solved
                print(f"CAPTCHA still processing... (Attempt {attempt}/{max_attempts})")
                continue
            
            else:
                # Unknown status
                raise Exception(f"Unknown task status: {result.get('status')}")
        
        # If we've exceeded max attempts
        raise Exception("Timed out waiting for CAPTCHA solution")