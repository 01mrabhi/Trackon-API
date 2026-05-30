import base64
import time
import logging
import requests

logger = logging.getLogger(__name__)

def solve_captcha(image_bytes: bytes, api_key: str) -> str:
    """
    Submits a base64 encoded captcha image to 2Captcha API
    and polls for the resolved text solution.
    """
    if not api_key:
        raise ValueError("2Captcha API key is required but missing.")
        
    try:
        # 1. Base64 encode the image
        img_base64 = base64.b64encode(image_bytes).decode("utf-8")
        
        # 2. Upload to 2Captcha
        logger.info("Uploading CAPTCHA to 2Captcha...")
        upload_url = "http://2captcha.com/in.php"
        payload = {
            "key": api_key,
            "method": "base64",
            "body": img_base64,
            "json": 1
        }
        
        response = requests.post(upload_url, json=payload, timeout=10)
        res_data = response.json()
        
        if res_data.get("status") != 1:
            error_msg = res_data.get("request", "Unknown upload error")
            raise RuntimeError(f"2Captcha upload failed: {error_msg}")
            
        captcha_id = res_data.get("request")
        logger.info(f"CAPTCHA uploaded successfully. ID: {captcha_id}. Polling for result...")
        
        # 3. Poll for the solution
        poll_url = "http://2captcha.com/res.php"
        params = {
            "key": api_key,
            "action": "get",
            "id": captcha_id,
            "json": 1
        }
        
        start_time = time.time()
        # Poll up to 60 seconds (30 iterations of 2s)
        for attempt in range(30):
            time.sleep(2)
            logger.info(f"Polling 2Captcha (attempt {attempt + 1})...")
            poll_resp = requests.get(poll_url, params=params, timeout=10)
            poll_data = poll_resp.json()
            
            status = poll_data.get("status")
            request_text = poll_data.get("request")
            
            if status == 1:
                logger.info(f"CAPTCHA solved successfully! Solution: '{request_text}' (took {int(time.time() - start_time)}s)")
                return str(request_text).strip()
                
            if request_text != "CAPCHA_NOT_READY":
                raise RuntimeError(f"2Captcha solving error: {request_text}")
                
        raise TimeoutError("2Captcha solving timed out (60s limit reached).")
        
    except Exception as e:
        logger.error(f"Error in solve_captcha: {e}")
        raise
