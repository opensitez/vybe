# vybe-test: python/while_loops/while_truthy_nonempty_string_iterates_chars
# origin: languages/python/tests/python/test_while_loops.rs

s = 'ab'
i = 0
while i < len(s):
 print(s[i])
 i += 1
