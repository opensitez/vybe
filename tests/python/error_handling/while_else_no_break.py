# vybe-test: python/error_handling/while_else_no_break
# origin: languages/python/tests/python/test_error_handling.rs

i = 0
while i < 3:
    i += 1
else:
    print('completed')
