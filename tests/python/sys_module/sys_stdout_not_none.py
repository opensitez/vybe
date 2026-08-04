# vybe-test: python/sys_module/sys_stdout_not_none
# origin: languages/python/tests/python/test_sys_module.rs

import sys
print(sys.stdout is not None)
