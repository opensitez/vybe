# vybe-test: python/import_system/importlib_metadata_distributions
# origin: languages/python/tests/python/test_import_system.rs

try:
 import importlib.metadata as md
 print(hasattr(md, 'version'))
except ImportError:
 print('skip')
