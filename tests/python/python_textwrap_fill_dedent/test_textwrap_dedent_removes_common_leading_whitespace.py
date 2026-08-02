# vybe-test: python/python_textwrap_fill_dedent/test_textwrap_dedent_removes_common_leading_whitespace
# origin: languages/python/tests/python/test_python_textwrap_fill_dedent.rs

import textwrap
text = "    line1\n    line2\n    line3"
print(textwrap.dedent(text))
