# vybe-test: python/augmented_assign/augmented_concat_in_loop
# origin: languages/python/tests/python/test_augmented_assign.rs

s = ''
for ch in 'ab':
 s += ch
print(s)
