# vybe-test: python/python_filecmp_comparison/test_filecmp_cmp_binary_content_matches
# origin: languages/python/tests/python/test_python_filecmp_comparison.rs

import filecmp, tempfile, os, shutil
d = tempfile.mkdtemp()
a = os.path.join(d, "a.bin")
b = os.path.join(d, "b.bin")
for p in [a, b]:
    with open(p, "wb") as f:
        f.write(bytes(range(256)))
print(filecmp.cmp(a, b, shallow=False))
shutil.rmtree(d)
