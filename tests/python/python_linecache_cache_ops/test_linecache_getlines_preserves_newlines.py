# vybe-test: python/python_linecache_cache_ops/test_linecache_getlines_preserves_newlines
# origin: languages/python/tests/python/test_python_linecache_cache_ops.rs

import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("x\ny\n")
f.close()
linecache.clearcache()
lines = linecache.getlines(f.name)
print(all(l.endswith("\n") for l in lines))
os.unlink(f.name)
