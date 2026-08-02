# vybe-test: python/while_loops/while_string_builder_concat
# origin: languages/python/tests/python/test_while_loops.rs

parts = ['a', 'b', 'c']
i = 0
s = ''
while i < len(parts):
 s += parts[i]
 i += 1
print(s)
