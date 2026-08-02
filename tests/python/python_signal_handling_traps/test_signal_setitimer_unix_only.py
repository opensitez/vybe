# vybe-test: python/python_signal_handling_traps/test_signal_setitimer_unix_only
# origin: languages/python/tests/python/test_python_signal_handling_traps.rs

import signal, sys
if hasattr(signal, "setitimer"):
    old = signal.setitimer(signal.ITIMER_REAL, 0)
    print(isinstance(old, tuple))
else:
    print(True)
