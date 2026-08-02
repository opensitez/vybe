# vybe-test: python/context_managers_extended/contextlib_redirect_stdout
# origin: languages/python/tests/python/test_context_managers_extended.rs

from contextlib import redirect_stdout
import io
buf = io.StringIO()
with redirect_stdout(buf):
 print('hidden')
print(len(buf.getvalue()) > 0)
