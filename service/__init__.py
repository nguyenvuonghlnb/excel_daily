import requests
import telegram
access_token = ""


def get_new_token():
    payload = 'grant_type=password&client_id=FiinTek.DataFeed.Client&client_secret=datafeed1212&scope=openid%20fiintek.datafeed&username=df.dsc%40fiingroup.vn&password=fkosKLzc(09F'
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    response = requests.request("POST", "http://identity.fiingroup.vn/connect/token", headers=headers, data=payload)
    if response.status_code == 200:
        body = response.json()
        return body["access_token"]
    return ""


def get_headers(force_update=False):
    global access_token
    if force_update:
        access_token = get_new_token()
    headers = {
        'Authorization': 'Bearer ' + access_token
    }
    return headers


def make_request(url, data={}):
    headers = get_headers()
    response = requests.request("GET", url, headers=headers, data=data)
    if response.status_code == 200:
        print(f"Success {response.status_code}: ok")
        return response.json()
    elif response.status_code == 401:
        get_headers(True)
        print("Error 401: retry get new token")
        return make_request(url, data)
    telegram.send(mess=f"Error {response.status_code}: {response.content}")
    print(f"Error {response.status_code}: {response.text}")
    return None
