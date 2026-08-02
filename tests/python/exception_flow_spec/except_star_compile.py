# vybe-test: python/exception_flow_spec/except_star_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

try:
    raise ExceptionGroup('group', [ValueError('a')])
except* ValueError:
    pass
