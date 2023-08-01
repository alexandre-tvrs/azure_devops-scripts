import requests
import settings
from requests.auth import HTTPBasicAuth


header = {"Content-Type": "application/json", "Accept": "application/json"}
auth = HTTPBasicAuth('', settings.token)


def InvokeREST(uri, method = 'get', data = []):
    if method in ['put', 'post', 'patch']:
        if len(data) == 0:
            response = 'Error: Insert data in function'
        else:
            response = (requests.request(method=method, url=uri, json=data, auth=auth, headers=header)).json()
    else:
        response = (requests.request(method=method, url=uri, auth=auth, headers=header)).json()

    return response

