use super::helpers::run_python;

// threading — Thread, Lock, RLock, Semaphore, BoundedSemaphore, Event, Condition, Barrier, local, current_thread, main_thread, active_count

#[test]
fn test_threading_thread_start_and_join() {
    let out = run_python(
        r#"
import threading

result = []
def worker(val):
    result.append(val * 2)

t = threading.Thread(target=worker, args=(21,))
t.start()
t.join()
print(result)
"#,
    );
    assert_eq!(out, vec!["[42]"]);
}

#[test]
fn test_threading_lock_acquire_release() {
    let out = run_python(
        r#"
import threading
lock = threading.Lock()
acquired = lock.acquire(blocking=False)
print(acquired)
print(lock.locked())
lock.release()
print(lock.locked())
"#,
    );
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_threading_lock_context_manager() {
    let out = run_python(
        r#"
import threading
lock = threading.Lock()
with lock:
    print(lock.locked())
print(lock.locked())
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_threading_rlock_reentrant() {
    let out = run_python(
        r#"
import threading
rlock = threading.RLock()
with rlock:
    with rlock:
        print("re-acquired successfully")
"#,
    );
    assert_eq!(out, vec!["re-acquired successfully"]);
}

#[test]
fn test_threading_semaphore_acquire_release() {
    let out = run_python(
        r#"
import threading
sem = threading.Semaphore(2)
print(sem.acquire())
print(sem.acquire())
print(sem.acquire(blocking=False))  # should fail (3rd acquire when max 2)
sem.release()
print(sem.acquire(blocking=False))  # now succeeds
"#,
    );
    assert_eq!(out, vec!["True", "True", "False", "True"]);
}

#[test]
fn test_threading_bounded_semaphore_raises_value_error_on_extra_release() {
    let out = run_python(
        r#"
import threading
bsem = threading.BoundedSemaphore(1)
bsem.acquire()
bsem.release()
try:
    bsem.release()
except ValueError:
    print("ValueError")
"#,
    );
    assert_eq!(out, vec!["ValueError"]);
}

#[test]
fn test_threading_event_set_wait_clear() {
    let out = run_python(
        r#"
import threading
evt = threading.Event()
print(evt.is_set())
evt.set()
print(evt.is_set())
print(evt.wait(timeout=0.01))
evt.clear()
print(evt.is_set())
"#,
    );
    assert_eq!(out, vec!["False", "True", "True", "False"]);
}

#[test]
fn test_threading_condition_notify_all() {
    let out = run_python(
        r#"
import threading
cond = threading.Condition()
state = []

def worker():
    with cond:
        state.append("ready")
        cond.notify()

with cond:
    t = threading.Thread(target=worker)
    t.start()
    cond.wait(timeout=1.0)
    print(state)
t.join()
"#,
    );
    assert_eq!(out, vec!["['ready']"]);
}

#[test]
fn test_threading_thread_local_isolation() {
    let out = run_python(
        r#"
import threading
local_data = threading.local()
local_data.val = "main_thread_val"

def worker():
    local_data.val = "worker_thread_val"

t = threading.Thread(target=worker)
t.start()
t.join()
print(local_data.val)
"#,
    );
    assert_eq!(out, vec!["main_thread_val"]);
}

#[test]
fn test_threading_current_thread_name() {
    let out = run_python(
        r#"
import threading
t = threading.current_thread()
print(isinstance(t.name, str))
print(len(t.name) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_threading_main_thread_check() {
    let out = run_python(
        r#"
import threading
print(threading.current_thread() is threading.main_thread())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_threading_active_count() {
    let out = run_python(
        r#"
import threading
print(threading.active_count() >= 1)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_threading_thread_is_alive() {
    let out = run_python(
        r#"
import threading, time

def long_running():
    time.sleep(0.05)

t = threading.Thread(target=long_running)
print(t.is_alive())
t.start()
print(t.is_alive())
t.join()
print(t.is_alive())
"#,
    );
    assert_eq!(out, vec!["False", "True", "False"]);
}

#[test]
fn test_threading_thread_daemon_attribute() {
    let out = run_python(
        r#"
import threading
t = threading.Thread(daemon=True)
print(t.daemon)
t.daemon = False
print(t.daemon)
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_threading_barrier_wait() {
    let out = run_python(
        r#"
import threading
barrier = threading.Barrier(2)
passed = []

def worker():
    barrier.wait()
    passed.append(1)

t = threading.Thread(target=worker)
t.start()
barrier.wait()
t.join()
print(len(passed))
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_threading_timer_execution() {
    let out = run_python(
        r#"
import threading, time
flag = []
def set_flag():
    flag.append(True)

timer = threading.Timer(0.01, set_flag)
timer.start()
timer.join()
print(flag)
"#,
    );
    assert_eq!(out, vec!["[True]"]);
}

#[test]
fn test_threading_timer_cancel() {
    let out = run_python(
        r#"
import threading, time
flag = []
def set_flag():
    flag.append(True)

timer = threading.Timer(0.1, set_flag)
timer.start()
timer.cancel()
time.sleep(0.15)
print(flag)
"#,
    );
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_threading_get_ident() {
    let out = run_python(
        r#"
import threading
ident = threading.get_ident()
print(isinstance(ident, int))
print(ident > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_threading_enumerate_threads() {
    let out = run_python(
        r#"
import threading
threads = threading.enumerate()
print(len(threads) >= 1)
print(threading.main_thread() in threads)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_threading_lock_acquire_timeout() {
    let out = run_python(
        r#"
import threading
lock = threading.Lock()
lock.acquire()
# Second acquire with timeout should fail
res = lock.acquire(timeout=0.01)
print(res)
lock.release()
"#,
    );
    assert_eq!(out, vec!["False"]);
}
