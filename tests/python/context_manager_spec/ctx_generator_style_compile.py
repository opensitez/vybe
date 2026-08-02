# vybe-test: python/context_manager_spec/ctx_generator_style_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

from contextlib import contextmanager
@contextmanager
def cm():
    yield 1
