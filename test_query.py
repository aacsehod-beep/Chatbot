import requests
res = requests.post("http://127.0.0.1:5000/ask", json={"message": "Where is the library?"})
print(res.json()['response'])
