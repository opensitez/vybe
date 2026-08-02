# vybe-test: python/python_os_scandir_direntry/test_os_get_terminal_size_fallback
# origin: languages/python/tests/python/test_python_os_scandir_direntry.rs

import os
try:
    ts = os.get_terminal_size(0)
    print(isinstance(ts.columns, int))
except (OSError, ValueError):
    print("OSError")
