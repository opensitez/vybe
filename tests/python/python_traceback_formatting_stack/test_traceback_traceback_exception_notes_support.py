# vybe-test: python/python_traceback_formatting_stack/test_traceback_traceback_exception_notes_support
# origin: languages/python/tests/python/test_python_traceback_formatting_stack.rs

import traceback, sys
if sys.version_info >= (3, 11):
    try:
        e = ValueError("orig error")
        e.add_note("Note 1: check config")
        raise e
    except ValueError as exc:
        te = traceback.TracebackException.from_exception(exc)
        formatted = "".join(te.format())
        print("Note 1: check config" in formatted)
else:
    print(True)
