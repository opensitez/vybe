# vybe-test: python/python_walrus_scopes/test_walrus_in_while
# origin: languages/python/tests/python/test_python_walrus_scopes.rs

import io
stream = io.StringIO("line1\nline2\nline3\n")
lines = []
while line := stream.readline():
    lines.append(line.strip())
print(lines)
