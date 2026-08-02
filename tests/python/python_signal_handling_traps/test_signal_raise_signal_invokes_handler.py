# vybe-test: python/python_signal_handling_traps/test_signal_raise_signal_invokes_handler
# origin: languages/python/tests/python/test_python_signal_handling_traps.rs

import signal, sys
if hasattr(signal, "raise_signal"):
    catches = []
    def my_h(signum, frame):
        catches.append(signum)
    old = signal.signal(signal.SIGUSR1 if hasattr(signal, "SIGUSR1") else signal.SIGINT, my_h)
    target_sig = signal.SIGUSR1 if hasattr(signal, "SIGUSR1") else signal.SIGINT
    signal.raise_signal(target_sig)
    print(len(catches) == 1)
    signal.signal(target_sig, old)
else:
    print(True)
