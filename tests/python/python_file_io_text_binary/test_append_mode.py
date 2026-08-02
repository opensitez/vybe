# vybe-test: python/python_file_io_text_binary/test_append_mode
# origin: languages/python/tests/python/test_python_file_io_text_binary.rs

import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    path = f.name
    f.write("first\n")

with open(path, 'a') as f:
    f.write("second\n")

with open(path) as f:
    lines = [l.strip() for l in f.readlines()]

os.unlink(path)
print(lines)
