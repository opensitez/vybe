# vybe-test: python/os_path_extended/os_path_scandir
# origin: languages/python/tests/python/test_os_path_extended.rs
# vybe-test-mode: compile

import os
list(os.scandir('.'))
