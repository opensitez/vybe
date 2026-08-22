# vybe-test: python/functools_itertools/functools_cached_property
# origin: languages/python/tests/python/test_functools_itertools.rs

from functools import cached_property
class C:
 @cached_property
 def x(self):
  return 1
