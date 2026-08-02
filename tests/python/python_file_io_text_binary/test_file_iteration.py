# vybe-test: python/python_file_io_text_binary/test_file_iteration
# origin: languages/python/tests/python/test_python_file_io_text_binary.rs

import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    path = f.name
    f.write("a\nb\nc\n")

lines = []
with open(path) as f:
    for line in f:
        lines.append(line.strip())

os.unlink(path)
print(lines)
