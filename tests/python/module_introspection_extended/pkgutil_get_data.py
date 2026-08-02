# vybe-test: python/module_introspection_extended/pkgutil_get_data
# origin: languages/python/tests/python/test_module_introspection_extended.rs
# vybe-test-mode: compile

import pkgutil
pkgutil.get_data('json', 'decoder.py')
