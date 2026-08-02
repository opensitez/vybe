# vybe-test: python/python_sysconfig_paths/test_sysconfig_paths_all_strings
# origin: languages/python/tests/python/test_python_sysconfig_paths.rs

import sysconfig
paths = sysconfig.get_paths()
print(all(isinstance(v, str) for v in paths.values()))
