# vybe-test: python/python_textwrap_fill_dedent/test_textwrap_fill_preserves_paragraph_structure
# origin: languages/python/tests/python/test_python_textwrap_fill_dedent.rs

import textwrap
text = "aaa bbb ccc ddd eee"
result = textwrap.fill(text, width=10)
for line in result.splitlines():
    print(len(line) <= 10)
