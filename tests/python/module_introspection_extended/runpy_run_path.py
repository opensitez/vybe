# vybe-test: python/module_introspection_extended/runpy_run_path
# origin: languages/python/tests/python/test_module_introspection_extended.rs
# vybe-test-mode: compile

import runpy
runpy.run_path('.', run_name='__main__')
