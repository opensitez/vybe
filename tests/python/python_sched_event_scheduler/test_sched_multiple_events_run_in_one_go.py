# vybe-test: python/python_sched_event_scheduler/test_sched_multiple_events_run_in_one_go
# origin: languages/python/tests/python/test_python_sched_event_scheduler.rs

import sched

vtime = 0
s = sched.scheduler(lambda: vtime, lambda d: None)
executed = []

for i in range(5):
    s.enter(0, i, executed.append, (i,))

s.run()
print(executed)
