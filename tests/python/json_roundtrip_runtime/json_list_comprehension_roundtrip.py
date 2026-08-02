# vybe-test: python/json_roundtrip_runtime/json_list_comprehension_roundtrip
# origin: languages/python/tests/python/test_json_roundtrip_runtime.rs

import json
vals = [json.loads(json.dumps(x)) for x in [1, 'a', None]]
print(vals)
