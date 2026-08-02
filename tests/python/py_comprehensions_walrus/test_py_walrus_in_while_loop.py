# vybe-test: python/py_comprehensions_walrus/test_py_walrus_in_while_loop
# origin: languages/python/tests/python/test_py_comprehensions_walrus.rs

import io

buf = io.StringIO("line1\nline2\nline3\n")
lines = []
while (line := buf.readline()):
    lines.append(line.strip())
print(lines)
