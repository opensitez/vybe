# vybe-test: python/python_site_directory_configuration/test_site_getsitepackages_with_custom_prefixes
# origin: languages/python/tests/python/test_python_site_directory_configuration.rs

import site
prefixes = ['/custom/prefix1', '/custom/prefix2']
res = site.getsitepackages(prefixes)
print(any('/custom/prefix1' in p for p in res))
print(any('/custom/prefix2' in p for p in res))
