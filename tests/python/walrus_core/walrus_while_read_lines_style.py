# vybe-test: python/walrus_core/walrus_while_read_lines_style
# origin: languages/python/tests/python/test_walrus_core.rs

lines = iter(['x', ''])
out = []
while (line := next(lines, '')):
 out.append(line)
print(out)
