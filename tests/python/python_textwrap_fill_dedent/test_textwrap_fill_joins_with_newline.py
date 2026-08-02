# vybe-test: python/python_textwrap_fill_dedent/test_textwrap_fill_joins_with_newline
# origin: languages/python/tests/python/test_python_textwrap_fill_dedent.rs

import textwrap
result = textwrap.fill("one two three four five six", width=15)
print(result)
