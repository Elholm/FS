import requests as req

# print(requests.__version__)
# print(requests.__copyright__)

# resp = req.request(method='GET', url="http://esdh/")
# print(resp.text)

for i in range(1):
  user = 'jael'
  passwd = 'Juli20!9'

# resp = req.get("http://esdh/", auth=(user, passwd))
import requests
url = 'http://esdh/'
values = {'username': user,
          'password': passwd}

r = requests.post(url, data=values)
print(r.content)
print(r.status_code)
# print(resp.text)
# print(resp.status_code)