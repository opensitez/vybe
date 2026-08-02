# vybe-test: python/new_features/open_readline
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

f = open('test.txt')
line = f.readline()
f.close()
