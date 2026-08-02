# vybe-test: python/python_pydoc_documentation_renderer/test_pydoc_locate_error_handling
# origin: languages/python/tests/python/test_python_pydoc_documentation_renderer.rs

import pydoc
try:
    obj = pydoc.locate("non_existent_module_9999.invalid")
    print(obj is None)
except Exception:
    print(True)
