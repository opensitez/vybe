# vybe-test: python/python_mmap_file_io/test_mmap_anon_mmap_unix
# origin: languages/python/tests/python/test_python_mmap_file_io.rs

import mmap, sys
if sys.platform != "win32":
    mm = mmap.mmap(-1, 256)  # anonymous mmap
    mm.write(b"test data!")
    mm.seek(0)
    print(mm.read(10))
    mm.close()
else:
    print(b"test data!")
