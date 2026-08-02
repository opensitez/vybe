# vybe-test: python/python_signal_handling_traps/test_signal_alarm_unix_only
# origin: languages/python/tests/python/test_python_signal_handling_traps.rs

import signal, sys
if hasattr(signal, "alarm"):
    old_alarm = signal.alarm(0)
    print(isinstance(old_alarm, int))
else:
    print(True)
