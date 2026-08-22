# vybe-test: python/closure_truthiness_spec/bool_from_len_compile
# origin: languages/python/tests/python/test_closure_truthiness_spec.rs

class Box:
    def __len__(self):
        return 1
