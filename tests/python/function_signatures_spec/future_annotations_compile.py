# vybe-test: python/function_signatures_spec/future_annotations_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs

from __future__ import annotations
def f(x: 'Node') -> 'Node':
    return x
