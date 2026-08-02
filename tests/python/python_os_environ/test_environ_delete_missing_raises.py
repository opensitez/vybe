# vybe-test: python/python_os_environ/test_environ_delete_missing_raises
# origin: languages/python/tests/python/test_python_os_environ.rs

import os
try:
    del os.environ['__NO_SUCH_VAR_VYBE__']
    print("no_error")
except KeyError:
    print("KeyError")
