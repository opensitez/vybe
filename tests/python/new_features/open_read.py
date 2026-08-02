# vybe-test: python/new_features/open_read
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

f = open('test.txt', 'r')
data = f.read()
f.close()
