# vybe-test: python/py_zipfile_tarfile/test_py_zipfile_infolist_iteration
# origin: languages/python/tests/python/test_py_zipfile_tarfile.rs

import zipfile, io

buf = io.BytesIO()
with zipfile.ZipFile(buf, "w") as zf:
    for name in ["file1.txt", "file2.txt", "file3.txt"]:
        zf.writestr(name, f"content of {name}")

buf.seek(0)
with zipfile.ZipFile(buf, "r") as zf:
    filenames = [info.filename for info in zf.infolist()]
    print(filenames)
