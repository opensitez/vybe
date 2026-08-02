# vybe-test: python/error_handling/for_else_no_break
# origin: languages/python/tests/python/test_error_handling.rs

for x in [1, 2, 3]:
    if x == 5:
        break
else:
    print('no five')
