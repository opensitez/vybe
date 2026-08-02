# vybe-test: python/py_builtins_adv/test_py_builtins_iter_and_next
# origin: languages/python/tests/python/test_py_builtins_adv.rs

it = iter([1, 2, 3, 4])
print(next(it))
print(next(it))
print(list(it))  # consume rest

# iter with sentinel
import io
buf = io.StringIO("a\nb\nc\n")
lines = list(iter(buf.readline, ""))
print([l.strip() for l in lines])
