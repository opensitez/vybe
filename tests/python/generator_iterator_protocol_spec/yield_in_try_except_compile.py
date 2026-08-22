# vybe-test: python/generator_iterator_protocol_spec/yield_in_try_except_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs

def gen():
    try:
        yield 1
    except Exception:
        yield 2
