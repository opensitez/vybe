# vybe-test: python/python_filecmp_comparison/test_filecmp_dircmp_subdirs
# origin: languages/python/tests/python/test_python_filecmp_comparison.rs

import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
for d in [d1, d2]:
    os.makedirs(os.path.join(d, "sub"))
dc = filecmp.dircmp(d1, d2)
print("sub" in dc.subdirs)
shutil.rmtree(d1); shutil.rmtree(d2)
