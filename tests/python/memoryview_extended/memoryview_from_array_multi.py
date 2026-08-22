# vybe-test: python/memoryview_extended/memoryview_from_array_multi
# origin: languages/python/tests/python/test_memoryview_extended.rs

import array
a = array.array('d', [1.0, 2.0])
memoryview(a)
