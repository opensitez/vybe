# vybe-test: python/python_signal_handling_traps/test_signal_strsignal_description
# origin: languages/python/tests/python/test_python_signal_handling_traps.rs

import signal, sys
if hasattr(signal, "strsignal"):
    desc = signal.strsignal(signal.SIGINT)
    print(isinstance(desc, str))
    print(len(desc) > 0)
else:
    print(True)
    print(True)
