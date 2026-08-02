# vybe-test: python/while_loops/while_parses_digits_from_string
# origin: languages/python/tests/python/test_while_loops.rs

s = '1234'
i = 0
n = 0
while i < len(s):
 n = n * 10 + int(s[i])
 i += 1
print(n)
