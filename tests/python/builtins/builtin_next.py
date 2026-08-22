# vybe-test: python/builtins/builtin_next
# origin: languages/python/tests/python/test_builtins.rs
# `next()` needs an ITERATOR; a list is only iterable.
items = [1, 2, 3]
x = next(iter(items))
