# vybe-test: python/python_filecmp_comparison/test_filecmp_dircmp_same_files
# origin: languages/python/tests/python/test_python_filecmp_comparison.rs

import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
for d in [d1, d2]:
    with open(os.path.join(d, "same.txt"), "w") as f: f.write("identical")
dc = filecmp.dircmp(d1, d2)
print(dc.same_files)
shutil.rmtree(d1); shutil.rmtree(d2)
