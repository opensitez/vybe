# vybe-test: python/dict_comprehension/dict_comp_from_split
# origin: languages/python/tests/python/test_dict_comprehension.rs

{p: len(p) for p in 'a,b'.split(',')}
