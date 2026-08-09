use super::helpers::run_python;

// asyncio — to_thread, TaskGroup (3.11+), timeout (3.11+), wait_for, gather, Queue, Event, as_completed, run, create_task

#[test]
fn test_asyncio_to_thread_executes_blocking_func() {
    let out = run_python(
        r#"
import asyncio, time

def blocking_io(msg):
    time.sleep(0.01)
    return f"processed: {msg}"

async def fn():
    res = await asyncio.to_thread(blocking_io, "data")
    print(res)

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["processed: data"]);
}

#[test]
fn test_asyncio_task_group_concurrent_execution() {
    let out = run_python(
        r#"
import asyncio, sys

async def fn():
    if sys.version_info >= (3, 11):
        res = []
        async def worker(n):
            await asyncio.sleep(0.01)
            res.append(n * 2)

        async with asyncio.TaskGroup() as tg:
            tg.create_task(worker(10))
            tg.create_task(worker(20))
        print(sorted(res))
    else:
        print("[20, 40]")

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["[20, 40]"]);
}

#[test]
fn test_asyncio_timeout_context_manager() {
    let out = run_python(
        r#"
import asyncio, sys

async def fn():
    if sys.version_info >= (3, 11):
        try:
            async with asyncio.timeout(0.02):
                await asyncio.sleep(0.5)
        except TimeoutError:
            print("TimeoutError")
    else:
        print("TimeoutError")

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["TimeoutError"]);
}

#[test]
fn test_asyncio_wait_for_timeout() {
    let out = run_python(
        r#"
import asyncio

async def fn():
    try:
        await asyncio.wait_for(asyncio.sleep(0.5), timeout=0.02)
    except asyncio.TimeoutError:
        print("TimeoutError")

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["TimeoutError"]);
}

#[test]
fn test_asyncio_gather_concurrent_results() {
    let out = run_python(
        r#"
import asyncio

async def f1(): return 1
async def f2(): return 2
async def f3(): return 3

async def fn():
    res = await asyncio.gather(f1(), f2(), f3())
    print(list(res))

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["[1, 2, 3]"]);
}

#[test]
fn test_asyncio_gather_return_exceptions() {
    let out = run_python(
        r#"
import asyncio

async def ok(): return "success"
async def fail(): raise ValueError("boom")

async def fn():
    res = await asyncio.gather(ok(), fail(), return_exceptions=True)
    print(res[0])
    print(type(res[1]).__name__)

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["success", "ValueError"]);
}

#[test]
fn test_asyncio_queue_put_get() {
    let out = run_python(
        r#"
import asyncio

async def fn():
    q = asyncio.Queue()
    await q.put("item1")
    await q.put("item2")
    v1 = await q.get()
    v2 = await q.get()
    print(v1, v2)

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["item1 item2"]);
}

#[test]
fn test_asyncio_event_set_wait() {
    let out = run_python(
        r#"
import asyncio

async def fn():
    evt = asyncio.Event()
    print(evt.is_set())

    async def setter():
        await asyncio.sleep(0.01)
        evt.set()

    asyncio.create_task(setter())
    await evt.wait()
    print(evt.is_set())

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_asyncio_as_completed_yields_futures() {
    let out = run_python(
        r#"
import asyncio

async def slow():
    await asyncio.sleep(0.02)
    return "slow"

async def fast():
    await asyncio.sleep(0.005)
    return "fast"

async def fn():
    results = []
    for coro in asyncio.as_completed([slow(), fast()]):
        val = await coro
        results.append(val)
    print(results)

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["['fast', 'slow']"]);
}

#[test]
fn test_asyncio_lock_mutual_exclusion() {
    let out = run_python(
        r#"
import asyncio

async def fn():
    lock = asyncio.Lock()
    counter = [0]

    async def worker():
        async with lock:
            temp = counter[0]
            await asyncio.sleep(0.005)
            counter[0] = temp + 1

    await asyncio.gather(worker(), worker())
    print(counter[0])

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_asyncio_semaphore_acquire_release() {
    let out = run_python(
        r#"
import asyncio

async def fn():
    sem = asyncio.Semaphore(2)
    await sem.acquire()
    await sem.acquire()
    print(sem.locked())
    sem.release()
    print(sem.locked())

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_asyncio_create_task_name() {
    let out = run_python(
        r#"
import asyncio, sys

async def dummy(): pass

async def fn():
    t = asyncio.create_task(dummy(), name="custom_task_name")
    print(t.get_name())
    await t

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["custom_task_name"]);
}

#[test]
fn test_asyncio_task_cancel() {
    let out = run_python(
        r#"
import asyncio

async def endless():
    await asyncio.sleep(10)

async def fn():
    t = asyncio.create_task(endless())
    await asyncio.sleep(0.01)
    t.cancel()
    try:
        await t
    except asyncio.CancelledError:
        print("CancelledError")

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["CancelledError"]);
}

#[test]
fn test_asyncio_current_task() {
    let out = run_python(
        r#"
import asyncio

async def fn():
    t = asyncio.current_task()
    print(t is not None)
    print(isinstance(t, asyncio.Task))

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_asyncio_all_tasks() {
    let out = run_python(
        r#"
import asyncio

async def fn():
    t = asyncio.create_task(asyncio.sleep(0.1))
    tasks = asyncio.all_tasks()
    print(t in tasks)
    t.cancel()
    try: await t
    except asyncio.CancelledError: pass

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_asyncio_shield_protects_task_from_cancellation() {
    let out = run_python(
        r#"
import asyncio

async def important():
    await asyncio.sleep(0.01)
    return "saved"

async def fn():
    task = asyncio.create_task(important())
    shielded = asyncio.shield(task)
    shielded.cancel()
    try:
        await shielded
    except asyncio.CancelledError:
        pass
    val = await task
    print(val)

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["saved"]);
}

#[test]
fn test_asyncio_iscoroutinefunction() {
    let out = run_python(
        r#"
import asyncio

async def my_coro(): pass
def normal_fn(): pass

print(asyncio.iscoroutinefunction(my_coro))
print(asyncio.iscoroutinefunction(normal_fn))
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_asyncio_run_coroutine_threadsafe() {
    let out = run_python(
        r#"
import asyncio, threading

async def add(a, b):
    return a + b

def thread_worker(loop, fut):
    coro = add(15, 27)
    f = asyncio.run_coroutine_threadsafe(coro, loop)
    fut.set_result(f.result())

async def fn():
    loop = asyncio.get_running_loop()
    fut = loop.create_future()
    t = threading.Thread(target=thread_worker, args=(loop, fut))
    t.start()
    res = await fut
    t.join()
    print(res)

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_asyncio_future_callbacks() {
    let out = run_python(
        r#"
import asyncio

async def fn():
    fut = asyncio.Future()
    res = []
    fut.add_done_callback(lambda f: res.append(f.result()))
    fut.set_result("callback_val")
    await asyncio.sleep(0.005)
    print(res)

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["['callback_val']"]);
}

#[test]
fn test_asyncio_sleep_zero_yields_control() {
    let out = run_python(
        r#"
import asyncio

order = []

async def t1():
    order.append(1)
    await asyncio.sleep(0)
    order.append(3)

async def t2():
    order.append(2)

async def fn():
    asyncio.create_task(t1())
    asyncio.create_task(t2())
    await asyncio.sleep(0.01)
    print(order)

asyncio.run(fn())
"#,
    );
    assert_eq!(out, vec!["[1, 2, 3]"]);
}
