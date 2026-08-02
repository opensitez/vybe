# vybe-test: python/module_introspection_extended/module_getattr_missing
# origin: languages/python/tests/python/test_module_introspection_extended.rs

import json
try:
 json.no_attr_xyz
 print('ok')
except AttributeError:
 print('err')
