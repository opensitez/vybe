# vybe-test: python/exception_flow_spec/exitstack_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs

from contextlib import ExitStack
with ExitStack() as stack:
    pass
