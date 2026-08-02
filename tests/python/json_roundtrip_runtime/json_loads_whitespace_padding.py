# vybe-test: python/json_roundtrip_runtime/json_loads_whitespace_padding
# origin: languages/python/tests/python/test_json_roundtrip_runtime.rs

import json; json.loads('  [1]  ')
