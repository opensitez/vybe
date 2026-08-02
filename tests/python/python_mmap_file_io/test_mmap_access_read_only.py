# vybe-test: python/python_mmap_file_io/test_mmap_access_read_only
# origin: languages/python/tests/python/test_python_mmap_file_io.rs

import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"readonly data")
f.flush()
f.close()
with open(f.name, "rb") as fh:
    mm = mmap.mmap(fh.fileno(), 0, access=mmap.ACCESS_READ)
    print(mm[:8])
    try:
        mm[0:4] = b"XXXX"
        print("no error")
    except TypeError:
        print("TypeError")
    mm.close()
os.unlink(f.name)
