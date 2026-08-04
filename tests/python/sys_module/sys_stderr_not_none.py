# vybe-test: python/sys_module/sys_stderr_not_none
# origin: languages/python/tests/python/test_sys_module.rs

import sys
print(sys.stderr is not None)
