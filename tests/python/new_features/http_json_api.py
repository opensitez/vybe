# vybe-test: python/new_features/http_json_api
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

import requests
import json
response = requests.get('https://api.example.com/data')
data = json.loads(response)
print(data)
