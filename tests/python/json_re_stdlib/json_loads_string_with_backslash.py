# vybe-test: python/json_re_stdlib/json_loads_string_with_backslash
# origin: languages/python/tests/python/test_json_re_stdlib.rs

import json
print(json.loads('"a\\nb"'))
