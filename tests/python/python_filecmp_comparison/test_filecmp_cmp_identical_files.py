# vybe-test: python/python_filecmp_comparison/test_filecmp_cmp_identical_files
# origin: languages/python/tests/python/test_python_filecmp_comparison.rs

import filecmp, tempfile, os
d = tempfile.mkdtemp()
a = os.path.join(d, "a.txt")
b = os.path.join(d, "b.txt")
for p in [a, b]:
    with open(p, "w") as f:
        f.write("same content")
print(filecmp.cmp(a, b, shallow=False))
import shutil; shutil.rmtree(d)
