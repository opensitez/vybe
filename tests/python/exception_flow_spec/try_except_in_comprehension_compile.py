# vybe-test: python/exception_flow_spec/try_except_in_comprehension_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs

def f(xs):
    out = []
    for x in xs:
        try:
            out.append(x)
        except Exception:
            pass
    return out
