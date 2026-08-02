# vybe-test: python/for_while_extended/while_else_continue_path
# origin: languages/python/tests/python/test_for_while_extended.rs

i = 0
while i < 2:
 i += 1
 continue
else:
 print('else')
