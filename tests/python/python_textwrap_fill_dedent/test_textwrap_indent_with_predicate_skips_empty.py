# vybe-test: python/python_textwrap_fill_dedent/test_textwrap_indent_with_predicate_skips_empty
# origin: languages/python/tests/python/test_python_textwrap_fill_dedent.rs

import textwrap
result = textwrap.indent("line1\n\nline2", prefix="> ", predicate=lambda s: s.strip())
print(result)
