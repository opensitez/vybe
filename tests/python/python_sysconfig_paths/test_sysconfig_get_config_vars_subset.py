# vybe-test: python/python_sysconfig_paths/test_sysconfig_get_config_vars_subset
# origin: languages/python/tests/python/test_python_sysconfig_paths.rs

import sysconfig
result = sysconfig.get_config_vars("py_version", "prefix")
print(len(result) == 2)
print(all(isinstance(v, str) for v in result))
