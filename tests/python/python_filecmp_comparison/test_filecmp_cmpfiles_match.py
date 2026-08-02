# vybe-test: python/python_filecmp_comparison/test_filecmp_cmpfiles_match
# origin: languages/python/tests/python/test_python_filecmp_comparison.rs

import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
for d in [d1, d2]:
    with open(os.path.join(d, "x.txt"), "w") as f:
        f.write("identical")
match, mismatch, errors = filecmp.cmpfiles(d1, d2, ["x.txt"], shallow=False)
print(match)
print(mismatch)
shutil.rmtree(d1); shutil.rmtree(d2)
