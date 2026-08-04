# vybe-test: python/py_subprocess_process_management/test_py_subprocess_popen_pipe_communication
# origin: languages/python/tests/python/test_py_subprocess_process_management.rs

import subprocess

proc = subprocess.Popen(
    ["python3", "-c", "import sys; data = sys.stdin.read(); sys.stdout.write(f'REPLY:{data}')"],
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    text=True
)
stdout, _ = proc.communicate(input="REQUEST_DATA")
print(stdout)
print(proc.returncode)
