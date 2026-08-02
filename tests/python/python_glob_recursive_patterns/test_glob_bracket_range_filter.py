# vybe-test: python/python_glob_recursive_patterns/test_glob_bracket_range_filter
# origin: languages/python/tests/python/test_python_glob_recursive_patterns.rs

import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
for c in "abc123":
    open(os.path.join(d, f"x{c}.txt"), "w").close()
results = glob.glob(os.path.join(d, "x[a-c].txt"))
print(len(results))
shutil.rmtree(d)
