# vybe-test: python/collections_extended/ordereddict_popitem
# origin: languages/python/tests/python/test_collections_extended.rs

from collections import OrderedDict
d = OrderedDict(a=1, b=2)
d.popitem(last=False)
