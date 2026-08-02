# vybe-test: python/functools_itertools/functools_partialmethod
# origin: languages/python/tests/python/test_functools_itertools.rs
# vybe-test-mode: compile

from functools import partialmethod
class C:
 def m(self, x):
  return x
 f = partialmethod(m, 1)
