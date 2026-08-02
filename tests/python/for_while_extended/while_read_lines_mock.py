# vybe-test: python/for_while_extended/while_read_lines_mock
# origin: languages/python/tests/python/test_for_while_extended.rs

lines = ['a', 'b']
i = 0
while i < len(lines):
 print(lines[i])
 i += 1
 break
