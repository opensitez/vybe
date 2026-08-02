# vybe-test: python/python_glob_recursive_patterns/test_glob_root_dir_kwarg
# origin: languages/python/tests/python/test_python_glob_recursive_patterns.rs

import glob, tempfile, os, shutil, sys
d = tempfile.mkdtemp()
open(os.path.join(d, "target.py"), "w").close()
# root_dir added in 3.10
if sys.version_info >= (3, 10):
    results = glob.glob("*.py", root_dir=d)
    print(results)
else:
    print(["target.py"])
shutil.rmtree(d)
