# vybe-test: python/py_json_serialization_custom/test_py_json_primitive_types_roundtrip
# origin: languages/python/tests/python/test_py_json_serialization_custom.rs

import json

primitives = [None, True, False, 100, 3.14, "hello string", [1, 2], {"k": "v"}]
for p in primitives:
    s = json.dumps(p)
    restored = json.loads(s)
    print(restored == p)
