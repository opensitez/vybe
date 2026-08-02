# vybe-test: python/import_system/from_import_attr_error
# origin: languages/python/tests/python/test_import_system.rs

try:
 from json import missing_attr_xyz
 print('ok')
except ImportError:
 print('err')
