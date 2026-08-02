# vybe-test: python/python_traceback_formatting_stack/test_traceback_traceback_exception_max_group_depth
# origin: languages/python/tests/python/test_python_traceback_formatting_stack.rs

import traceback, sys
if sys.version_info >= (3, 11):
    eg = ExceptionGroup("group", [ValueError(1), TypeError(2)])
    te = traceback.TracebackException.from_exception(eg)
    formatted = "".join(te.format())
    print("ExceptionGroup: group" in formatted)
else:
    print(True)
