# vybe-test: python/python_subprocess_execution_pipes/test_subprocess_stdout_stderr_redirect_stderr_to_stdout
# origin: languages/python/tests/python/test_python_subprocess_execution_pipes.rs

import subprocess, sys
code = "import sys; sys.stdout.write('out\\n'); sys.stderr.write('err\\n')"
proc = subprocess.Popen([sys.executable, "-c", code], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
stdout, _ = proc.communicate()
print("out" in stdout and "err" in stdout)
