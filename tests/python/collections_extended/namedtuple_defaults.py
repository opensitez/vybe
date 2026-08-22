# vybe-test: python/collections_extended/namedtuple_defaults
# origin: languages/python/tests/python/test_collections_extended.rs

from collections import namedtuple
P = namedtuple('P', 'x y', defaults=[0])
P(1)
