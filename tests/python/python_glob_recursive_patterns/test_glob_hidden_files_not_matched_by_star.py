# vybe-test: python/python_glob_recursive_patterns/test_glob_hidden_files_not_matched_by_star
# origin: languages/python/tests/python/test_python_glob_recursive_patterns.rs

import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
open(os.path.join(d, ".hidden"), "w").close()
open(os.path.join(d, "visible"), "w").close()
results = glob.glob(os.path.join(d, "*"))
names = [os.path.basename(r) for r in results]
print("visible" in names)
print(".hidden" not in names)
shutil.rmtree(d)
