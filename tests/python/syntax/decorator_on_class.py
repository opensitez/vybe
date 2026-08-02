# vybe-test: python/syntax/decorator_on_class
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

@dataclass
class Point:
    x: int
    y: int
