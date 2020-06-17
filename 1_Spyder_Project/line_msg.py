import requests

def Line_MSG(token:str, msg:str):
    headers = {"Authorization":"Bearer " + token
               ,"Content-Type":"application/x-www-form-urlencoded"}
    payload = {'message': msg}
    r = requests.post("https://notify-api.line.me/api/notify",
                      headers=headers, params=payload)
    return r.status_code


