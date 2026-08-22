# vybe-test: python/context_manager_spec/ctx_attr_manager_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# The base/name this fixture uses was never defined — supplied so it RUNS.
import contextlib as _c, types as _t
@_c.contextmanager
def _mgr():
    yield 1
resource_holder = _t.SimpleNamespace(manager=_mgr)


obj = resource_holder
with obj.manager() as r:
    pass
