# vybe-test: python/operators_arithmetic/len_uses_user_dunder_len
# origin: languages/python/tests/python/test_operators_arithmetic.rs

class C:
    def __len__(s): return 7
len(C())
