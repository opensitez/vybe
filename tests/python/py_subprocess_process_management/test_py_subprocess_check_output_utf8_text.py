# vybe-test: python/py_subprocess_process_management/test_py_subprocess_check_output_utf8_text
# origin: languages/python/tests/python/test_py_subprocess_process_management.rs

import subprocess

out = subprocess.check_output(["python3", "-c", "import sys; sys.stdout.write('output_string')"], text=True)
print(out)
