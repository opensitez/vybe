# vybe-test: python/python_linecache_cache_ops/test_linecache_getlines_count_matches_file
# origin: languages/python/tests/python/test_python_linecache_cache_ops.rs

import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
for i in range(10):
    f.write(f"line {i}\n")
f.close()
linecache.clearcache()
lines = linecache.getlines(f.name)
print(len(lines))
os.unlink(f.name)
