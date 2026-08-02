# vybe-test: python/import_system/zipimport_module
# origin: languages/python/tests/python/test_import_system.rs

try:
 import zipimport
 print(hasattr(zipimport, 'zipimporter'))
except ImportError:
 print('skip')
