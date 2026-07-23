use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Threading & Synchronization Primitives — Thread, Lock, RLock, Event, Condition, Semaphore, Barrier, local, current_thread
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_threading_lock_counter_protection() {
    let src = r#"
import threading

counter = 0
lock = threading.Lock()

def worker():
    global counter
    for _ in range(100):
        with lock:
            counter += 1

threads = [threading.Thread(target=worker) for _ in range(5)]
for t in threads: t.start()
for t in threads: t.join()

print(counter)
"#;
    assert_eq!(run_python(src), vec!["500"]);
}

#[test]
fn test_py_threading_rlock_reentrant_nesting() {
    let src = r#"
import threading

rlock = threading.RLock()
log = []

def outer():
    with rlock:
        log.append("outer")
        inner()

def inner():
    with rlock:
        log.append("inner")

outer()
print(log)
"#;
    assert_eq!(run_python(src), vec!["['outer', 'inner']"]);
}

#[test]
fn test_py_threading_event_signaling_wait() {
    let src = r#"
import threading

evt = threading.Event()
out = []

def waiter():
    evt.wait()
    out.append("notified")

t = threading.Thread(target=waiter)
t.start()
out.append("setting_event")
evt.set()
t.join()

print(out)
"#;
    assert_eq!(run_python(src), vec!["['setting_event', 'notified']"]);
}

#[test]
fn test_py_threading_condition_wait_notify() {
    let src = r#"
import threading

cond = threading.Condition()
state = []

def consumer():
    with cond:
        cond.wait()
        state.append("consumed")

def producer():
    with cond:
        state.append("produced")
        cond.notify()

c = threading.Thread(target=consumer)
p = threading.Thread(target=producer)
c.start()
import time; time.sleep(0.01)
p.start()
c.join()
p.join()
print(state)
"#;
    assert_eq!(run_python(src), vec!["['produced', 'consumed']"]);
}

#[test]
fn test_py_threading_semaphore_max_concurrency() {
    let src = r#"
import threading

sem = threading.Semaphore(2)
active = 0
max_active = 0
lock = threading.Lock()

def worker():
    global active, max_active
    with sem:
        with lock:
            active += 1
            if active > max_active: max_active = active
        import time; time.sleep(0.005)
        with lock:
            active -= 1

threads = [threading.Thread(target=worker) for _ in range(4)]
for t in threads: t.start()
for t in threads: t.join()

print(max_active <= 2)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_threading_barrier_synchronization() {
    let src = r#"
import threading

barrier = threading.Barrier(3)
log = []

def worker(idx):
    log.append(f"w{idx}_before")
    barrier.wait()
    log.append(f"w{idx}_after")

threads = [threading.Thread(target=worker, args=(i,)) for i in range(3)]
for t in threads: t.start()
for t in threads: t.join()

print(len([x for x in log if "after" in x]) == 3)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_threading_local_thread_storage_isolation() {
    let src = r#"
import threading

local_data = threading.local()
out = {}

def worker(val):
    local_data.val = val
    import time; time.sleep(0.005)
    out[threading.current_thread().name] = local_data.val

t1 = threading.Thread(target=worker, args=(10,), name="T1")
t2 = threading.Thread(target=worker, args=(20,), name="T2")
t1.start(); t2.start()
t1.join(); t2.join()

print(out["T1"], out["T2"])
"#;
    assert_eq!(run_python(src), vec!["10 20"]);
}

#[test]
fn test_py_threading_current_main_thread_check() {
    let src = r#"
import threading

main_t = threading.main_thread()
curr_t = threading.current_thread()
print(curr_t is main_t)
print(curr_t.is_alive())
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_threading_daemon_thread_flag() {
    let src = r#"
import threading

t = threading.Thread(target=lambda: None, daemon=True)
print(t.daemon)
t.start()
t.join()
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_threading_timer_delayed_execution() {
    let src = r#"
import threading

out = []

def delayed_task():
    out.append("timer_executed")

timer = threading.Timer(0.01, delayed_task)
timer.start()
timer.join()
print(out)
"#;
    assert_eq!(run_python(src), vec!["['timer_executed']"]);
}
