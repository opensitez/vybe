# vybe-test: python/function_signatures_spec/future_annotations_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# vybe-test-mode: compile

from __future__ import annotations
def f(x: 'Node') -> 'Node':
    return x
