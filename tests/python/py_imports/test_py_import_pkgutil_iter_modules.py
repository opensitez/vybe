# vybe-test: python/py_imports/test_py_import_pkgutil_iter_modules
# origin: languages/python/tests/python/test_py_imports.rs

import pkgutil, sys

# Check that standard library modules are discoverable
names = [m.name for m in pkgutil.iter_modules() if m.name in ("json", "math", "os")]
print(sorted(set(names)))
