# vybe-test: python/syntax/annotation_dict_type
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

def foo(x: dict[str, list[int]]) -> None:
    pass
