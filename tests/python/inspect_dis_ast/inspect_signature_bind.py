# vybe-test: python/inspect_dis_ast/inspect_signature_bind
# origin: languages/python/tests/python/test_inspect_dis_ast.rs
# vybe-test-mode: compile

import inspect
def f(a, *, b): pass
inspect.signature(f).bind(1, b=2)
