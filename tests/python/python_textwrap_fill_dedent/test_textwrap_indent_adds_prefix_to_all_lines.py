# vybe-test: python/python_textwrap_fill_dedent/test_textwrap_indent_adds_prefix_to_all_lines
# origin: languages/python/tests/python/test_python_textwrap_fill_dedent.rs

import textwrap
result = textwrap.indent("line1\nline2", prefix="  ")
print(result)
