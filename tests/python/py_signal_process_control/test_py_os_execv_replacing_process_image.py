# vybe-test: python/py_signal_process_control/test_py_os_execv_replacing_process_image
# origin: languages/python/tests/python/test_py_signal_process_control.rs

import os, sys

pid = os.fork()
if pid == 0:
    # Child replaces image
    os.execv(sys.executable, [sys.executable, "-c", "print('execv_success')"])
else:
    _, status = os.waitpid(pid, 0)
    print(os.WEXITSTATUS(status) == 0)
