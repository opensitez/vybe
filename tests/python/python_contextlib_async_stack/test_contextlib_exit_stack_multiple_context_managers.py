# vybe-test: python/python_contextlib_async_stack/test_contextlib_exit_stack_multiple_context_managers
# origin: languages/python/tests/python/test_python_contextlib_async_stack.rs

import contextlib, io

buf1 = io.StringIO()
buf2 = io.StringIO()

with contextlib.ExitStack() as stack:
    f1 = stack.enter_context(contextlib.redirect_stdout(buf1))
    f2 = stack.enter_context(contextlib.redirect_stderr(buf2))
    print("to stdout")

print(buf1.getvalue().strip())
