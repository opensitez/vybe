# vybe-test: python/collections_extended/chainmap_new_child
# origin: languages/python/tests/python/test_collections_extended.rs
# vybe-test-mode: compile

from collections import ChainMap
cm = ChainMap({'a': 1})
cm.new_child({'b': 2})
