# vybe-test: python/python_site_directory_configuration/test_site_getsitepackages_contains_expected_directory_structure
# origin: languages/python/tests/python/test_python_site_directory_configuration.rs

import site, sys
paths = site.getsitepackages()
print(any(sys.prefix in p or sys.exec_prefix in p for p in paths))
