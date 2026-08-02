# vybe-test: python/new_features/open_readlines
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

f = open('test.txt')
lines = f.readlines()
f.close()
