# vybe-test: python/import_system/runpy_run_path
# origin: languages/python/tests/python/test_import_system.rs
# vybe-test-mode: compile

import runpy
runpy.run_path('.', run_name='__main__')
