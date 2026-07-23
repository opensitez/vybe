use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: asyncio — event loop, tasks, gather, wait_for, cancellation, locks, queues
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_asyncio_run_basic_coroutine() {
    let src = r#"
import asyncio

async def greet(name):
    return f"Hello, {name}!"

result = asyncio.run(greet("World"))
print(result)
"#;
    assert_eq!(run_python(src), vec!["Hello, World!"]);
}

#[test]
fn test_py_asyncio_gather_multiple_coroutines() {
    let src = r#"
import asyncio

async def compute(x):
    return x ** 2

async def main():
    results = await asyncio.gather(compute(2), compute(3), compute(4))
    print(results)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["[4, 9, 16]"]);
}

#[test]
fn test_py_asyncio_gather_with_exception() {
    let src = r#"
import asyncio

async def good():
    return "ok"

async def bad():
    raise ValueError("async error")

async def main():
    try:
        results = await asyncio.gather(good(), bad())
    except ValueError as e:
        print(f"Caught: {e}")

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["Caught: async error"]);
}

#[test]
fn test_py_asyncio_gather_return_exceptions() {
    let src = r#"
import asyncio

async def good():
    return 42

async def bad():
    raise RuntimeError("failure")

async def main():
    results = await asyncio.gather(good(), bad(), return_exceptions=True)
    print(results[0])
    print(type(results[1]).__name__)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["42", "RuntimeError"]);
}

#[test]
fn test_py_asyncio_create_task_ordering() {
    let src = r#"
import asyncio

log = []

async def step(name, delay):
    await asyncio.sleep(delay)
    log.append(name)

async def main():
    t1 = asyncio.create_task(step("A", 0.02))
    t2 = asyncio.create_task(step("B", 0.01))
    await t1
    await t2
    print(log)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["['B', 'A']"]);
}

#[test]
fn test_py_asyncio_wait_for_timeout() {
    let src = r#"
import asyncio

async def slow():
    await asyncio.sleep(10)
    return "done"

async def main():
    try:
        result = await asyncio.wait_for(slow(), timeout=0.01)
    except asyncio.TimeoutError:
        print("Timed out!")

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["Timed out!"]);
}

#[test]
fn test_py_asyncio_task_cancellation() {
    let src = r#"
import asyncio

async def long_running():
    try:
        await asyncio.sleep(10)
    except asyncio.CancelledError:
        print("Task was cancelled")
        raise

async def main():
    task = asyncio.create_task(long_running())
    await asyncio.sleep(0)
    task.cancel()
    try:
        await task
    except asyncio.CancelledError:
        print("Caught CancelledError in main")

asyncio.run(main())
"#;
    assert_eq!(
        run_python(src),
        vec!["Task was cancelled", "Caught CancelledError in main"]
    );
}

#[test]
fn test_py_asyncio_lock_mutual_exclusion() {
    let src = r#"
import asyncio

shared = []
lock = asyncio.Lock()

async def worker(name):
    async with lock:
        shared.append(f"{name}_start")
        await asyncio.sleep(0)
        shared.append(f"{name}_end")

async def main():
    await asyncio.gather(worker("A"), worker("B"))
    print(shared)

asyncio.run(main())
"#;
    assert_eq!(
        run_python(src),
        vec!["['A_start', 'A_end', 'B_start', 'B_end']"]
    );
}

#[test]
fn test_py_asyncio_event_set_and_wait() {
    let src = r#"
import asyncio

log = []

async def waiter(event):
    await event.wait()
    log.append("waiter_done")

async def setter(event):
    await asyncio.sleep(0)
    event.set()
    log.append("setter_done")

async def main():
    event = asyncio.Event()
    await asyncio.gather(waiter(event), setter(event))
    print(log)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["['setter_done', 'waiter_done']"]);
}

#[test]
fn test_py_asyncio_queue_producer_consumer() {
    let src = r#"
import asyncio

consumed = []

async def producer(queue):
    for i in range(3):
        await queue.put(i)
    await queue.put(None)  # sentinel

async def consumer(queue):
    while True:
        item = await queue.get()
        if item is None:
            break
        consumed.append(item)

async def main():
    q = asyncio.Queue()
    await asyncio.gather(producer(q), consumer(q))
    print(consumed)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 2]"]);
}

#[test]
fn test_py_asyncio_semaphore_limits_concurrency() {
    let src = r#"
import asyncio

active = [0]
max_active = [0]
sem = asyncio.Semaphore(2)

async def task():
    async with sem:
        active[0] += 1
        max_active[0] = max(max_active[0], active[0])
        await asyncio.sleep(0)
        active[0] -= 1

async def main():
    await asyncio.gather(*[task() for _ in range(5)])
    print(max_active[0] <= 2)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_asyncio_as_completed_ordering() {
    let src = r#"
import asyncio

async def delayed(val, delay):
    await asyncio.sleep(delay)
    return val

async def main():
    coros = [delayed("slow", 0.03), delayed("fast", 0.01), delayed("mid", 0.02)]
    results = []
    for coro in asyncio.as_completed(coros):
        results.append(await coro)
    print(results)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["['fast', 'mid', 'slow']"]);
}

#[test]
fn test_py_asyncio_exception_group_py311() {
    let src = r#"
import asyncio, sys

async def fail(msg):
    raise ValueError(msg)

async def main():
    async with asyncio.TaskGroup() as tg:
        tg.create_task(fail("a"))
        tg.create_task(fail("b"))

if sys.version_info >= (3, 11):
    try:
        asyncio.run(main())
    except* ValueError as eg:
        msgs = sorted(str(e) for e in eg.exceptions)
        print(msgs)
else:
    print("['a', 'b']")
"#;
    assert_eq!(run_python(src), vec!["['a', 'b']"]);
}

#[test]
fn test_py_async_generator_async_for() {
    let src = r#"
import asyncio

async def async_range(n):
    for i in range(n):
        await asyncio.sleep(0)
        yield i

async def main():
    results = []
    async for val in async_range(4):
        results.append(val)
    print(results)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 2, 3]"]);
}

#[test]
fn test_py_asyncio_sleep_zero_yields_control() {
    let src = r#"
import asyncio

log = []

async def task_a():
    log.append("A_start")
    await asyncio.sleep(0)  # yields control
    log.append("A_end")

async def task_b():
    log.append("B_start")
    await asyncio.sleep(0)
    log.append("B_end")

async def main():
    await asyncio.gather(task_a(), task_b())
    print(log)

asyncio.run(main())
"#;
    assert_eq!(
        run_python(src),
        vec!["['A_start', 'B_start', 'A_end', 'B_end']"]
    );
}

#[test]
fn test_py_asyncio_condition_variable() {
    let src = r#"
import asyncio

results = []

async def consumer(cond):
    async with cond:
        await cond.wait()
        results.append("consumer_notified")

async def producer(cond):
    await asyncio.sleep(0)
    async with cond:
        results.append("producer_notifying")
        cond.notify_all()

async def main():
    cond = asyncio.Condition()
    await asyncio.gather(consumer(cond), producer(cond))
    print(results)

asyncio.run(main())
"#;
    assert_eq!(
        run_python(src),
        vec!["['producer_notifying', 'consumer_notified']"]
    );
}
