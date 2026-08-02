# vybe-test: python/python_collections_counter/test_python_counter_total_is_supported_fallback
# origin: languages/python/tests/python/test_python_collections_counter.rs

import sys
from collections import Counter
c = Counter(a=2, b=3)
if hasattr(c, 'total'):
    print(c.total())
else:
    print(sum(c.values()) if sys.version_info >= (3, 0) else 0)
