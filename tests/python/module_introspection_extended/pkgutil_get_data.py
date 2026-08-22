# vybe-test: python/module_introspection_extended/pkgutil_get_data
# origin: languages/python/tests/python/test_module_introspection_extended.rs

import pkgutil
pkgutil.get_data('json', 'decoder.py')
