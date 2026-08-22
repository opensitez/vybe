# vybe-test: python/comprehension_walrus_spec/comp_with_attr_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# The base/name this fixture uses was never defined — supplied so it RUNS.
import types as _t
items = [_t.SimpleNamespace(value=1), _t.SimpleNamespace(value=2)]


xs = [obj.value for obj in items]
