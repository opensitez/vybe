use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Async Coroutines & Event Loop — async/await, coroutines, gather, create_task, run, task result
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_async_coroutine_execution() {
    let src = r#"
import asyncio

async def fetch_data():
    return "data"

result = asyncio.run(fetch_data())
print(result)
"#;
    assert_eq!(run_python(src), vec!["data"]);
}

#[test]
fn test_py_async_gather_parallel_execution() {
    let src = r#"
import asyncio

async def compute(x):
    await asyncio.sleep(0.001)
    return x * 2

async def main():
    results = await asyncio.gather(compute(1), compute(2), compute(3))
    print(results)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["[2, 4, 6]"]);
}

#[test]
fn test_py_async_create_task_background() {
    let src = r#"
import asyncio

async def background_worker(name):
    await asyncio.sleep(0.001)
    return f"Done: {name}"

async def main():
    t1 = asyncio.create_task(background_worker("task1"))
    t2 = asyncio.create_task(background_worker("task2"))
    r1 = await t1
    r2 = await t2
    print(r1)
    print(r2)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["Done: task1", "Done: task2"]);
}

#[test]
fn test_py_async_wait_for_timeout() {
    let src = r#"
import asyncio

async def slow_task():
    await asyncio.sleep(1.0)
    return "completed"

async def main():
    try:
        await asyncio.wait_for(slow_task(), timeout=0.01)
    except asyncio.TimeoutError:
        print("TimeoutError caught")

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["TimeoutError caught"]);
}

#[test]
fn test_py_async_generator_async_for_loop() {
    let src = r#"
import asyncio

async def async_numbers(n):
    for i in range(n):
        await asyncio.sleep(0.001)
        yield i

async def main():
    out = []
    async for num in async_numbers(3):
        out.append(num)
    print(out)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 2]"]);
}

#[test]
fn test_py_async_context_manager_protocol() {
    let src = r#"
import asyncio

class AsyncConn:
    async def __aenter__(self):
        print("CONNECTED")
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        print("DISCONNECTED")

async def main():
    async with AsyncConn() as conn:
        print("QUERYING")

asyncio.run(main())
"#;
    assert_eq!(
        run_python(src),
        vec!["CONNECTED", "QUERYING", "DISCONNECTED"]
    );
}

#[test]
fn test_py_async_gather_return_exceptions() {
    let src = r#"
import asyncio

async def ok(): return "OK"
async def fail(): raise ValueError("FAIL")

async def main():
    results = await asyncio.gather(ok(), fail(), return_exceptions=True)
    print(results[0])
    print(type(results[1]).__name__, str(results[1]))

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["OK", "ValueError FAIL"]);
}

#[test]
fn test_py_async_task_cancellation() {
    let src = r#"
import asyncio

async def infinite():
    try:
        await asyncio.sleep(10)
    except asyncio.CancelledError:
        print("Cancelled inside task")
        raise

async def main():
    task = asyncio.create_task(infinite())
    await asyncio.sleep(0.001)
    task.cancel()
    try:
        await task
    except asyncio.CancelledError:
        print("CancelledError in main")

asyncio.run(main())
"#;
    assert_eq!(
        run_python(src),
        vec!["Cancelled inside task", "CancelledError in main"]
    );
}

#[test]
fn test_py_async_event_signaling() {
    let src = r#"
import asyncio

async def waiter(event):
    await event.wait()
    print("Event received")

async def main():
    ev = asyncio.Event()
    t = asyncio.create_task(waiter(ev))
    await asyncio.sleep(0.001)
    ev.set()
    await t

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["Event received"]);
}

#[test]
fn test_py_async_lock_concurrency_protection() {
    let src = r#"
import asyncio

counter = 0
lock = asyncio.Lock()

async def worker():
    global counter
    async with lock:
        val = counter
        await asyncio.sleep(0.001)
        counter = val + 1

async def main():
    await asyncio.gather(worker(), worker(), worker())
    print(counter)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["3"]);
}
