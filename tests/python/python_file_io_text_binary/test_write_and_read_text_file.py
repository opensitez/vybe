# vybe-test: python/python_file_io_text_binary/test_write_and_read_text_file
# origin: languages/python/tests/python/test_python_file_io_text_binary.rs

import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    path = f.name
    f.write("hello\nworld\n")

with open(path, 'r') as f:
    content = f.read()

os.unlink(path)
print(content)
