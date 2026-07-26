// Python threading.Event, Condition, Barrier — synchronization primitives
use super::helpers::run_python;

#[test]
fn test_event_set_wait() {
    let script = r#"
import threading

event = threading.Event()
results = []

def waiter():
    event.wait(timeout=5)
    results.append("done")

t = threading.Thread(target=waiter)
t.start()
event.set()
t.join()
print(results)
"#;
    assert_eq!(run_python(script), vec!["['done']"]);
}

#[test]
fn test_event_is_set() {
    let script = r#"
import threading

event = threading.Event()
print(event.is_set())
event.set()
print(event.is_set())
event.clear()
print(event.is_set())
"#;
    assert_eq!(run_python(script), vec!["False", "True", "False"]);
}

#[test]
fn test_event_timeout_returns_false() {
    let script = r#"
import threading

event = threading.Event()
result = event.wait(timeout=0.01)
print(result)
"#;
    assert_eq!(run_python(script), vec!["False"]);
}

#[test]
fn test_condition_notify() {
    let script = r#"
import threading

cond = threading.Condition()
results = []

def worker():
    with cond:
        cond.wait()
        results.append("notified")

t = threading.Thread(target=worker)
t.start()
import time
time.sleep(0.05)
with cond:
    cond.notify()
t.join()
print(results)
"#;
    assert_eq!(run_python(script), vec!["['notified']"]);
}

#[test]
fn test_semaphore_acquire_release() {
    let script = r#"
import threading

sem = threading.Semaphore(2)
acquired = []

def task(n):
    if sem.acquire(timeout=1):
        acquired.append(n)
        sem.release()

threads = [threading.Thread(target=task, args=(i,)) for i in range(4)]
for t in threads:
    t.start()
for t in threads:
    t.join()

print(len(acquired) == 4)
"#;
    assert_eq!(run_python(script), vec!["True"]);
}
