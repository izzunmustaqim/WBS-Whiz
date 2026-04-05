"""API client for communicating with the Fujitsu Gemini AI service.

Pure functions for sending requests to the AI API.
No GUI dependencies — callers are responsible for error display.
"""

import requests

API_ENDPOINT = (
    "https://api.ai-service.global.fujitsu.com"
    "/ai-foundation/chat-ai/gemini/pro:generateContent"
)


def send_gemini_request(api_key: str, prompt: str) -> str:
    """Send a prompt to the Fujitsu Gemini API and return the text response.

    Args:
        api_key: The API key for authentication.
        prompt: The prompt text to send.

    Returns:
        The text content from the API response.

    Raises:
        requests.exceptions.RequestException: For HTTP or connection errors.
        json.JSONDecodeError: If the response body is not valid JSON.
        KeyError: If the expected response structure is missing.
    """
    headers = {
        "Content-type": "application/json",
        "api-key": api_key
    }

    data = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": prompt
                    }
                ]
            }
        ]
    }

    response = requests.post(API_ENDPOINT, headers=headers, json=data)
    response.raise_for_status()

    analysis_result = response.json()
    content = analysis_result['candidates'][0]['content']['parts'][0]['text']
    return content
