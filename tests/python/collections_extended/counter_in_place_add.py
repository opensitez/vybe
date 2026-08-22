# vybe-test: python/collections_extended/counter_in_place_add
# origin: languages/python/tests/python/test_collections_extended.rs

from collections import Counter
c = Counter(a=1)
c += Counter(a=2)
