# vybe-test: python/walrus_extended/walrus_with_statement
# origin: languages/python/tests/python/test_walrus_extended.rs

class CM:
 def __enter__(self): return self
 def __exit__(self, *a): pass
# `with X as (c := ...)` is a SyntaxError — the `as` target is a binding
# TARGET, not an expression. The walrus belongs in the context expression.
with (c := CM()) as _w:
 pass
