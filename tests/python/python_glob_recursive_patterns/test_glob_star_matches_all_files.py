# vybe-test: python/python_glob_recursive_patterns/test_glob_star_matches_all_files
# origin: languages/python/tests/python/test_python_glob_recursive_patterns.rs

import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
for name in ["a.txt", "b.csv", "c.py"]:
    open(os.path.join(d, name), "w").close()
results = glob.glob(os.path.join(d, "*"))
print(len(results))
shutil.rmtree(d)
