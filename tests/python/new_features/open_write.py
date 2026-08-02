# vybe-test: python/new_features/open_write
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

f = open('test.txt', 'w')
f.write('hello')
f.close()
