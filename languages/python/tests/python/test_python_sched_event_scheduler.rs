use super::helpers::run_python;

// sched — scheduler, enter, enterabs, cancel, run, empty, queue, Event tuple, timefunc, delayfunc

#[test]
fn test_sched_scheduler_basic_execution_order() {
    let out = run_python(r#"
import sched, time

executed = []
s = sched.scheduler(time.time, time.sleep)

def task(name):
    executed.append(name)

s.enter(0.01, 1, task, ("first",))
s.enter(0.02, 1, task, ("second",))
s.run()

print(executed)
"#);
    assert_eq!(out, vec!["['first', 'second']"]);
}

#[test]
fn test_sched_priority_ordering_for_same_time_events() {
    let out = run_python(r#"
import sched, time

executed = []
s = sched.scheduler(time.time, time.sleep)

def log(p): executed.append(p)

now = time.time() + 0.05
s.enterabs(now, 2, log, ("priority_2",))
s.enterabs(now, 1, log, ("priority_1",))
s.enterabs(now, 3, log, ("priority_3",))
s.run()

print(executed)
"#);
    assert_eq!(out, vec!["['priority_1', 'priority_2', 'priority_3']"]);
}

#[test]
fn test_sched_cancel_pending_event() {
    let out = run_python(r#"
import sched, time

executed = []
s = sched.scheduler(time.time, time.sleep)

e1 = s.enter(0.01, 1, executed.append, ("task1",))
e2 = s.enter(0.02, 1, executed.append, ("task2",))

s.cancel(e1)
s.run()

print(executed)
"#);
    assert_eq!(out, vec!["['task2']"]);
}

#[test]
fn test_sched_empty_property_check() {
    let out = run_python(r#"
import sched, time

s = sched.scheduler(time.time, time.sleep)
print(s.empty())

e = s.enter(1.0, 1, print, ("test",))
print(s.empty())

s.cancel(e)
print(s.empty())
"#);
    assert_eq!(out, vec!["True", "False", "True"]);
}

#[test]
fn test_sched_queue_property_returns_events() {
    let out = run_python(r#"
import sched, time

s = sched.scheduler(time.time, time.sleep)
e1 = s.enter(10, 1, print, ("a",))
e2 = s.enter(5, 2, print, ("b",))

q = s.queue
print(len(q))
print(q[0].argument)
"#);
    assert_eq!(out, vec!["2", "('b',)"]);
}

#[test]
fn test_sched_custom_virtual_time_and_delay_functions() {
    let out = run_python(r#"
import sched

vtime = 0
log = []

def timefunc():
    return vtime

def delayfunc(delay):
    global vtime
    vtime += delay

s = sched.scheduler(timefunc, delayfunc)
s.enter(10, 1, log.append, ("event_at_10",))
s.enter(25, 1, log.append, ("event_at_25",))
s.run()

print(log)
print(vtime)
"#);
    assert_eq!(out, vec!["['event_at_10', 'event_at_25']", "25"]);
}

#[test]
fn test_sched_cancel_non_existent_event_raises_valueerror() {
    let out = run_python(r#"
import sched, time

s = sched.scheduler(time.time, time.sleep)
e = s.enter(1, 1, print, ("dummy",))
s.cancel(e)
try:
    s.cancel(e)
except ValueError:
    print("ValueError")
"#);
    assert_eq!(out, vec!["ValueError"]);
}

#[test]
fn test_sched_kwargs_passing() {
    let out = run_python(r#"
import sched, time

captured = {}
s = sched.scheduler(time.time, time.sleep)

def handler(*, name, val):
    captured[name] = val

s.enter(0.01, 1, handler, kwargs={"name": "param", "val": 100})
s.run()

print(captured)
"#);
    assert_eq!(out, vec!["{'param': 100}"]);
}

#[test]
fn test_sched_event_tuple_attributes() {
    let out = run_python(r#"
import sched, time

s = sched.scheduler(time.time, time.sleep)
e = s.enter(5, 2, print, ("arg",), {"kw": 1})

print(hasattr(e, "time"))
print(hasattr(e, "priority"))
print(hasattr(e, "action"))
print(hasattr(e, "argument"))
print(hasattr(e, "kwargs"))
print(e.priority)
"#);
    assert_eq!(out, vec!["True", "True", "True", "True", "True", "2"]);
}

#[test]
fn test_sched_run_blocking_false_non_blocking_execution() {
    let out = run_python(r#"
import sched, time

executed = []
s = sched.scheduler(time.time, time.sleep)
s.enter(100, 1, executed.append, ("future",))
s.enter(0, 1, executed.append, ("now",))

s.run(blocking=False)
print(executed)
print(len(s.queue))
"#);
    assert_eq!(out, vec!["['now']", "1"]);
}

#[test]
fn test_sched_scheduling_event_from_within_event_handler() {
    let out = run_python(r#"
import sched

vtime = 0
def timefunc(): return vtime
def delayfunc(d): global vtime; vtime += d

executed = []
s = sched.scheduler(timefunc, delayfunc)

def step2(): executed.append("step2")

def step1():
    executed.append("step1")
    s.enter(5, 1, step2)

s.enter(5, 1, step1)
s.run()

print(executed)
print(vtime)
"#);
    assert_eq!(out, vec!["['step1', 'step2']", "10"]);
}

#[test]
fn test_sched_event_comparison() {
    let out = run_python(r#"
import sched, time

s = sched.scheduler(time.time, time.sleep)
e1 = s.enter(10, 1, print)
e2 = s.enter(20, 1, print)
print(e1 < e2)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sched_enterabs_with_past_timestamp() {
    let out = run_python(r#"
import sched, time

executed = []
s = sched.scheduler(time.time, time.sleep)
s.enterabs(time.time() - 100, 1, executed.append, ("past_event",))
s.run()

print(executed)
"#);
    assert_eq!(out, vec!["['past_event']"]);
}

#[test]
fn test_sched_multiple_events_run_in_one_go() {
    let out = run_python(r#"
import sched

vtime = 0
s = sched.scheduler(lambda: vtime, lambda d: None)
executed = []

for i in range(5):
    s.enter(0, i, executed.append, (i,))

s.run()
print(executed)
"#);
    assert_eq!(out, vec!["[0, 1, 2, 3, 4]"]);
}

#[test]
fn test_sched_action_exception_stops_scheduler_unless_caught() {
    let out = run_python(r#"
import sched

vtime = 0
s = sched.scheduler(lambda: vtime, lambda d: None)

def faulty():
    raise RuntimeError("task failed")

s.enter(0, 1, faulty)
try:
    s.run()
except RuntimeError:
    print("RuntimeError")
"#);
    assert_eq!(out, vec!["RuntimeError"]);
}

#[test]
fn test_sched_queue_immutability_view() {
    let out = run_python(r#"
import sched, time

s = sched.scheduler(time.time, time.sleep)
s.enter(5, 1, print)
q = s.queue
s.enter(10, 1, print)
print(len(q))
"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_sched_default_delayfunc_timefunc() {
    let out = run_python(r#"
import sched

s = sched.scheduler()
print(s.timefunc is not None)
print(s.delayfunc is not None)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_sched_enter_returns_event_instance() {
    let out = run_python(r#"
import sched, time

s = sched.scheduler(time.time, time.sleep)
e = s.enter(1, 1, print)
print(isinstance(e, sched.Event))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sched_cancel_first_element_in_queue() {
    let out = run_python(r#"
import sched

vtime = 0
s = sched.scheduler(lambda: vtime, lambda d: None)
executed = []

e1 = s.enter(10, 1, executed.append, (1,))
e2 = s.enter(20, 1, executed.append, (2,))

s.cancel(e1)
s.run()
print(executed)
"#);
    assert_eq!(out, vec!["[2]"]);
}

#[test]
fn test_sched_run_on_empty_scheduler_returns_immediately() {
    let out = run_python(r#"
import sched, time

s = sched.scheduler(time.time, time.sleep)
res = s.run()
print(res is None)
"#);
    assert_eq!(out, vec!["True"]);
}
