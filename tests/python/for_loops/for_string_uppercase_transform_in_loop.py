# vybe-test: python/for_loops/for_string_uppercase_transform_in_loop
# origin: languages/python/tests/python/test_for_loops.rs

result = ''
for ch in 'ab':
    result += ch.upper()
print(result)
