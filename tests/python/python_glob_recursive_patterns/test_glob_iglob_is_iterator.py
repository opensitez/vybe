# vybe-test: python/python_glob_recursive_patterns/test_glob_iglob_is_iterator
# origin: languages/python/tests/python/test_python_glob_recursive_patterns.rs

import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
for i in range(3):
    open(os.path.join(d, f"f{i}.txt"), "w").close()
it = glob.iglob(os.path.join(d, "*.txt"))
results = list(it)
print(len(results))
shutil.rmtree(d)
