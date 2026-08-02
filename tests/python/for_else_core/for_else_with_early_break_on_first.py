# vybe-test: python/for_else_core/for_else_with_early_break_on_first
# origin: languages/python/tests/python/test_for_else_core.rs

for x in [9]:
 print(x)
 break
else:
 print('else')
print('after')
