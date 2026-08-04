# vybe-test: python/py_subprocess/test_py_subprocess_popen_pipe
# origin: languages/python/tests/python/test_py_subprocess.rs

import subprocess

proc = subprocess.Popen(
    ["python3", "-c", "import sys; sys.stdout.write('from_child')"],
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE
)
stdout, stderr = proc.communicate()
print(stdout.decode())
print(proc.returncode)
