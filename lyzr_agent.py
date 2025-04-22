import requests

def send_message_to_agent(user_id: str, agent_id: str, message: str, api_key: str) -> dict:
    url = 'https://agent-prod.studio.lyzr.ai/v3/inference/chat/'
    headers = {
        'Content-Type': 'application/json',
        'x-api-key': api_key
    }
    payload = {
        "user_id": user_id,
        "agent_id": agent_id,
        "session_id": agent_id,  # using agent_id as session_id as in the original curl
        "message": message
    }

    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        return {"error": str(e)}
