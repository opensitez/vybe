use super::helpers::run_python;

#[test]
fn test_fifo_queue_basic_put_get() {
    let out = run_python(r##"
import queue
q = queue.Queue()
q.put("first")
q.put("second")
print(q.get())
print(q.get())
"##);
    assert_eq!(out, vec!["first", "second"]);
}

#[test]
fn test_fifo_queue_maxsize_full() {
    let out = run_python(r##"
import queue
q = queue.Queue(maxsize=2)
print(q.empty())
q.put(10)
q.put(20)
print(q.full())
"##);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_queue_get_nowait_empty() {
    let out = run_python(r##"
import queue
q = queue.Queue()
try:
    q.get_nowait()
except queue.Empty:
    print("EMPTY_CAUGHT")
"##);
    assert_eq!(out, vec!["EMPTY_CAUGHT"]);
}

#[test]
fn test_queue_put_nowait_full() {
    let out = run_python(r##"
import queue
q = queue.Queue(maxsize=1)
q.put_nowait("a")
try:
    q.put_nowait("b")
except queue.Full:
    print("FULL_CAUGHT")
"##);
    assert_eq!(out, vec!["FULL_CAUGHT"]);
}

#[test]
fn test_lifo_queue_ordering() {
    let out = run_python(r##"
import queue
q = queue.LifoQueue()
q.put(1)
q.put(2)
q.put(3)
print(q.get())
print(q.get())
print(q.get())
"##);
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn test_priority_queue_tuple_priority() {
    let out = run_python(r##"
import queue
q = queue.PriorityQueue()
q.put((3, "low"))
q.put((1, "high"))
q.put((2, "medium"))
print(q.get()[1])
print(q.get()[1])
print(q.get()[1])
"##);
    assert_eq!(out, vec!["high", "medium", "low"]);
}

#[test]
fn test_priority_queue_custom_comparable() {
    let out = run_python(r##"
import queue
q = queue.PriorityQueue()
q.put((10, "task_b"))
q.put((5, "task_a"))
while not q.empty():
    priority, task = q.get()
    print(f"{task}:{priority}")
"##);
    assert_eq!(out, vec!["task_a:5", "task_b:10"]);
}

#[test]
fn test_simple_queue_unbounded() {
    let out = run_python(r##"
import queue
sq = queue.SimpleQueue()
sq.put("alpha")
sq.put("beta")
print(sq.get())
print(sq.get())
"##);
    assert_eq!(out, vec!["alpha", "beta"]);
}

#[test]
fn test_simple_queue_empty_exception() {
    let out = run_python(r##"
import queue
sq = queue.SimpleQueue()
try:
    sq.get_nowait()
except queue.Empty:
    print("SIMPLE_QUEUE_EMPTY")
"##);
    assert_eq!(out, vec!["SIMPLE_QUEUE_EMPTY"]);
}

#[test]
fn test_queue_task_done_join_completion() {
    let out = run_python(r##"
import queue
q = queue.Queue()
q.put("job1")
q.put("job2")
item1 = q.get()
q.task_done()
item2 = q.get()
q.task_done()
q.join()
print("JOIN_SUCCESSFUL")
"##);
    assert_eq!(out, vec!["JOIN_SUCCESSFUL"]);
}

#[test]
fn test_queue_task_done_excess_error() {
    let out = run_python(r##"
import queue
q = queue.Queue()
q.put("item")
q.get()
q.task_done()
try:
    q.task_done()
except ValueError:
    print("EXCESS_TASK_DONE")
"##);
    assert_eq!(out, vec!["EXCESS_TASK_DONE"]);
}

#[test]
fn test_queue_qsize_reporting() {
    let out = run_python(r##"
import queue
q = queue.Queue()
print(q.qsize())
q.put("x")
q.put("y")
print(q.qsize())
q.get()
print(q.qsize())
"##);
    assert_eq!(out, vec!["0", "2", "1"]);
}

#[test]
fn test_queue_with_mixed_types() {
    let out = run_python(r##"
import queue
q = queue.Queue()
q.put({"key": "value"})
q.put([1, 2, 3])
q.put((True, False))
print(q.get()["key"])
print(q.get()[2])
print(q.get()[0])
"##);
    assert_eq!(out, vec!["value", "3", "True"]);
}

#[test]
fn test_priority_queue_numeric_priorities() {
    let out = run_python(r##"
import queue
q = queue.PriorityQueue()
q.put((3.14, "pi"))
q.put((2.71, "e"))
q.put((1.41, "sqrt2"))
print(q.get()[1])
print(q.get()[1])
print(q.get()[1])
"##);
    assert_eq!(out, vec!["sqrt2", "e", "pi"]);
}

#[test]
fn test_lifo_queue_maxsize() {
    let out = run_python(r##"
import queue
q = queue.LifoQueue(maxsize=1)
q.put("first")
print(q.full())
try:
    q.put_nowait("second")
except queue.Full:
    print("LIFO_FULL")
"##);
    assert_eq!(out, vec!["True", "LIFO_FULL"]);
}

#[test]
fn test_queue_clear_via_drain() {
    let out = run_python(r##"
import queue
q = queue.Queue()
for i in range(5):
    q.put(i)
drained = []
while not q.empty():
    drained.append(q.get())
print(drained)
"##);
    assert_eq!(out, vec!["[0, 1, 2, 3, 4]"]);
}

#[test]
fn test_queue_timeout_parameter() {
    let out = run_python(r##"
import queue
q = queue.Queue()
try:
    q.get(timeout=0.01)
except queue.Empty:
    print("TIMEOUT_EMPTY")
"##);
    assert_eq!(out, vec!["TIMEOUT_EMPTY"]);
}

#[test]
fn test_simple_queue_qsize() {
    let out = run_python(r##"
import queue
sq = queue.SimpleQueue()
print(sq.empty())
sq.put("item1")
sq.put("item2")
print(sq.empty())
print(sq.qsize())
"##);
    assert_eq!(out, vec!["True", "False", "2"]);
}

#[test]
fn test_queue_put_block_false() {
    let out = run_python(r##"
import queue
q = queue.Queue(maxsize=1)
q.put("only_one")
try:
    q.put("overflow", block=False)
except queue.Full:
    print("BLOCK_FALSE_FULL")
"##);
    assert_eq!(out, vec!["BLOCK_FALSE_FULL"]);
}

#[test]
fn test_queue_get_block_false() {
    let out = run_python(r##"
import queue
q = queue.Queue()
try:
    q.get(block=False)
except queue.Empty:
    print("BLOCK_FALSE_EMPTY")
"##);
    assert_eq!(out, vec!["BLOCK_FALSE_EMPTY"]);
}
