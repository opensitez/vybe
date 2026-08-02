# vybe-test: python/pythonic_idioms/walrus_while_read_lines_pattern
# origin: languages/python/tests/python/test_pythonic_idioms.rs

lines = iter(['a', ''])
out = []
while (line := next(lines, '')):
 out.append(line)
print(out)
