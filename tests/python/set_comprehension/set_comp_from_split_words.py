# vybe-test: python/set_comprehension/set_comp_from_split_words
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({w for w in 'hi,ho'.split(',')})
