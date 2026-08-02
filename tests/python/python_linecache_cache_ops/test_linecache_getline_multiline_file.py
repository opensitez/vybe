# vybe-test: python/python_linecache_cache_ops/test_linecache_getline_multiline_file
# origin: languages/python/tests/python/test_python_linecache_cache_ops.rs

import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
lines = [f"row{i}\n" for i in range(5)]
f.write("".join(lines))
f.close()
linecache.clearcache()
for i, expected in enumerate(lines, start=1):
    got = linecache.getline(f.name, i)
    assert got == expected, f"line {i}: {got!r} != {expected!r}"
print("all ok")
os.unlink(f.name)
