# vybe-test: python/import_system/import_error_message
# origin: languages/python/tests/python/test_import_system.rs

try:
 import no_such_module_xyz_abc
 print('ok')
except ImportError:
 print('err')
