# vybe-test: python/python_glob_recursive_patterns/test_glob_iglob_lazy_vs_glob_list
# origin: languages/python/tests/python/test_python_glob_recursive_patterns.rs

import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
for i in range(4):
    open(os.path.join(d, f"x{i}.log"), "w").close()
pattern = os.path.join(d, "*.log")
print(sorted(glob.glob(pattern)) == sorted(glob.iglob(pattern)))
shutil.rmtree(d)
