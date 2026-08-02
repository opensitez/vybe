# vybe-test: python/python_sysconfig_paths/test_sysconfig_get_path_names
# origin: languages/python/tests/python/test_python_sysconfig_paths.rs

import sysconfig
names = sysconfig.get_path_names()
for name in ("stdlib", "platstdlib", "platlib", "purelib", "include", "scripts", "data"):
    print(name in names)
