# vybe-test: python/context_managers_core/with_redirect_stdout_style_buffer
# origin: languages/python/tests/python/test_context_managers_core.rs

from io import StringIO
from contextlib import redirect_stdout
buf = StringIO()
with redirect_stdout(buf):
 print('hidden')
print(buf.getvalue().strip())
