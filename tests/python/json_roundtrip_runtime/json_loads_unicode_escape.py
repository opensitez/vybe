# vybe-test: python/json_roundtrip_runtime/json_loads_unicode_escape
# origin: languages/python/tests/python/test_json_roundtrip_runtime.rs

import json; json.loads('"\\u00e9"')
