# vybe-test: python/python_reprlib_repr_limits/test_reprlib_maxdict_limits_dict_entries
# origin: languages/python/tests/python/test_python_reprlib_repr_limits.rs

import reprlib
r = reprlib.Repr()
r.maxdict = 2
d = {i: i*2 for i in range(10)}
result = r.repr(d)
print("..." in result)
