# vybe-test: python/functools_itertools/itertools_batched
# origin: languages/python/tests/python/test_functools_itertools.rs
# vybe-test-mode: compile

import itertools
list(itertools.batched([1,2,3], 2))
