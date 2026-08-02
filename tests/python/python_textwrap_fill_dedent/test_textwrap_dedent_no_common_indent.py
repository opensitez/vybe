# vybe-test: python/python_textwrap_fill_dedent/test_textwrap_dedent_no_common_indent
# origin: languages/python/tests/python/test_python_textwrap_fill_dedent.rs

import textwrap
text = "no\n  indent\nhere"
print(textwrap.dedent(text))
