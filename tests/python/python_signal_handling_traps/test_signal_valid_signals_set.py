# vybe-test: python/python_signal_handling_traps/test_signal_valid_signals_set
# origin: languages/python/tests/python/test_python_signal_handling_traps.rs

import signal, sys
if hasattr(signal, "valid_signals"):
    sigs = signal.valid_signals()
    print(signal.SIGINT in sigs)
    print(signal.SIGTERM in sigs)
else:
    print(True)
    print(True)
