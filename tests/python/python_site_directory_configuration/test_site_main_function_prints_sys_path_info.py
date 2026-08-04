# vybe-test: python/python_site_directory_configuration/test_site_main_function_prints_sys_path_info
# origin: languages/python/tests/python/test_python_site_directory_configuration.rs

import site, io, sys
buf = io.StringIO()
orig = sys.stdout
sys.stdout = buf
try:
    site.main()
finally:
    sys.stdout = orig
print('sys.path' in buf.getvalue())
