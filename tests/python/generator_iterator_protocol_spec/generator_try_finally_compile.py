# vybe-test: python/generator_iterator_protocol_spec/generator_try_finally_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

def gen():
    try:
        yield 1
    finally:
        cleanup()
