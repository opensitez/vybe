# vybe-test: python/import_system/importlib_import_module_reload
# origin: languages/python/tests/python/test_import_system.rs

import importlib
m = importlib.import_module('json')
importlib.reload(m)
