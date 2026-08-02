# vybe-test: python/python_signal_handling_traps/test_signal_sig_block_unblock_unix
# origin: languages/python/tests/python/test_python_signal_handling_traps.rs

import signal, sys
if hasattr(signal, "pthread_sigmask"):
    old_mask = signal.pthread_sigmask(signal.SIG_BLOCK, [])
    print(isinstance(old_mask, set))
else:
    print(True)
