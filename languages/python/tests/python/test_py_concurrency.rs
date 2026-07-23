use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: threading + multiprocessing — Thread, Lock, Queue, Pool, Process, shared memory
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_threading_basic_thread() {
    let src = r#"
import threading

results = []

def worker(n):
    results.append(n * n)

threads = [threading.Thread(target=worker, args=(i,)) for i in range(4)]
for t in threads:
    t.start()
for t in threads:
    t.join()

print(sorted(results))
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 4, 9]"]);
}

#[test]
fn test_py_threading_lock_mutual_exclusion() {
    let src = r#"
import threading

counter = [0]
lock = threading.Lock()

def increment():
    for _ in range(100):
        with lock:
            counter[0] += 1

threads = [threading.Thread(target=increment) for _ in range(5)]
for t in threads:
    t.start()
for t in threads:
    t.join()

print(counter[0])
"#;
    assert_eq!(run_python(src), vec!["500"]);
}

#[test]
fn test_py_threading_event_signaling() {
    let src = r#"
import threading

log = []
event = threading.Event()

def waiter():
    event.wait()
    log.append("received")

t = threading.Thread(target=waiter)
t.start()
log.append("firing")
event.set()
t.join()
print(log)
"#;
    assert_eq!(run_python(src), vec!["['firing', 'received']"]);
}

#[test]
fn test_py_threading_condition() {
    let src = r#"
import threading

log = []
cond = threading.Condition()

def consumer():
    with cond:
        cond.wait()
        log.append("consumed")

def producer():
    with cond:
        log.append("produced")
        cond.notify()

c = threading.Thread(target=consumer)
p = threading.Thread(target=producer)
c.start()
import time; time.sleep(0.01)
p.start()
c.join()
p.join()
print(log)
"#;
    assert_eq!(run_python(src), vec!["['produced', 'consumed']"]);
}

#[test]
fn test_py_threading_rlock_reentrant() {
    let src = r#"
import threading

rlock = threading.RLock()
results = []

def reentrant():
    with rlock:
        results.append("outer")
        with rlock:  # can reacquire
            results.append("inner")

t = threading.Thread(target=reentrant)
t.start()
t.join()
print(results)
"#;
    assert_eq!(run_python(src), vec!["['outer', 'inner']"]);
}

#[test]
fn test_py_threading_queue_producer_consumer() {
    let src = r#"
import threading
from queue import Queue

q = Queue()
results = []

def producer():
    for i in range(5):
        q.put(i)
    q.put(None)  # sentinel

def consumer():
    while True:
        item = q.get()
        if item is None:
            break
        results.append(item)

p = threading.Thread(target=producer)
c = threading.Thread(target=consumer)
p.start()
c.start()
p.join()
c.join()
print(results)
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 2, 3, 4]"]);
}

#[test]
fn test_py_threading_semaphore() {
    let src = r#"
import threading

sem = threading.Semaphore(2)
active = [0]
max_seen = [0]
lock = threading.Lock()

def task():
    with sem:
        with lock:
            active[0] += 1
            max_seen[0] = max(max_seen[0], active[0])
        import time; time.sleep(0.01)
        with lock:
            active[0] -= 1

threads = [threading.Thread(target=task) for _ in range(6)]
for t in threads:
    t.start()
for t in threads:
    t.join()

print(max_seen[0] <= 2)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_threading_daemon_thread() {
    let src = r#"
import threading

results = []

def bg():
    results.append("background")

t = threading.Thread(target=bg, daemon=True)
print(t.daemon)
t.start()
t.join()
print(results)
"#;
    assert_eq!(run_python(src), vec!["True", "['background']"]);
}

#[test]
fn test_py_threading_barrier() {
    let src = r#"
import threading

log = []
barrier = threading.Barrier(3)

def task(name):
    log.append(f"{name}_before")
    barrier.wait()
    log.append(f"{name}_after")

threads = [threading.Thread(target=task, args=(f"T{i}",)) for i in range(3)]
for t in threads:
    t.start()
for t in threads:
    t.join()

befores = sorted(x for x in log if "before" in x)
afters = sorted(x for x in log if "after" in x)
print(befores)
print(afters)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['T0_before', 'T1_before', 'T2_before']",
            "['T0_after', 'T1_after', 'T2_after']"
        ]
    );
}

#[test]
fn test_py_threading_local_storage() {
    let src = r#"
import threading

local = threading.local()
results = {}

def task(name, val):
    local.value = val
    import time; time.sleep(0.005)
    results[name] = local.value

threads = [
    threading.Thread(target=task, args=("A", 1)),
    threading.Thread(target=task, args=("B", 2)),
    threading.Thread(target=task, args=("C", 3)),
]
for t in threads:
    t.start()
for t in threads:
    t.join()

print(results["A"])
print(results["B"])
print(results["C"])
"#;
    assert_eq!(run_python(src), vec!["1", "2", "3"]);
}

#[test]
fn test_py_multiprocessing_pool_map() {
    let src = r#"
import multiprocessing

def square(x):
    return x ** 2

if __name__ == "__main__":
    with multiprocessing.Pool(2) as pool:
        results = pool.map(square, [1, 2, 3, 4, 5])
    print(results)
"#;
    assert_eq!(run_python(src), vec!["[1, 4, 9, 16, 25]"]);
}

#[test]
fn test_py_concurrent_futures_threadpool() {
    let src = r#"
from concurrent.futures import ThreadPoolExecutor

def compute(x):
    return x * x

with ThreadPoolExecutor(max_workers=4) as ex:
    futures = [ex.submit(compute, i) for i in range(5)]
    results = sorted(f.result() for f in futures)

print(results)
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 4, 9, 16]"]);
}

#[test]
fn test_py_concurrent_futures_processpool() {
    let src = r#"
from concurrent.futures import ProcessPoolExecutor

def square(n):
    return n * n

with ProcessPoolExecutor(max_workers=2) as ex:
    results = list(ex.map(square, range(5)))

print(results)
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 4, 9, 16]"]);
}

#[test]
fn test_py_concurrent_futures_as_completed() {
    let src = r#"
from concurrent.futures import ThreadPoolExecutor, as_completed
import time

def slow(n):
    time.sleep(n * 0.01)
    return n

with ThreadPoolExecutor(max_workers=4) as ex:
    futures = {ex.submit(slow, i): i for i in [3, 1, 2]}
    order = []
    for f in as_completed(futures):
        order.append(f.result())

print(sorted(order))
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]"]);
}
