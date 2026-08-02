# vybe-test: python/py_signal_process_control/test_py_os_strsignal_description
# origin: languages/python/tests/python/test_py_signal_process_control.rs

import signal, sys

if hasattr(signal, "strsignal"):
    desc = signal.strsignal(signal.SIGINT)
    print(isinstance(desc, str))
else:
    print("True")
