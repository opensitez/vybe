# vybe-test: python/while_loops/while_counts_vowels_in_string
# origin: languages/python/tests/python/test_while_loops.rs

s = 'hello'
i = 0
c = 0
while i < len(s):
 if s[i] in 'aeiou':
  c += 1
 i += 1
print(c)
