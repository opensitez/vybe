# vybe-test: python/python_glob_recursive_patterns/test_glob_extension_filter
# origin: languages/python/tests/python/test_python_glob_recursive_patterns.rs

import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
for name in ["a.py", "b.py", "c.txt", "d.rs"]:
    open(os.path.join(d, name), "w").close()
results = glob.glob(os.path.join(d, "*.py"))
print(len(results))
shutil.rmtree(d)
