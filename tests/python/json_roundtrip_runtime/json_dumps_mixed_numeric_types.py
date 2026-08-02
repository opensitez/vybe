# vybe-test: python/json_roundtrip_runtime/json_dumps_mixed_numeric_types
# origin: languages/python/tests/python/test_json_roundtrip_runtime.rs

import json; json.dumps([1, 2.5, -3])
