#![allow(non_snake_case)]
use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Async Coroutines, Tasks & Queues — asyncio.Queue, TaskGroup, Lock, Event, gather, wait_for
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_async_queue_producer_consumer() {
    let src = r#"
import asyncio

async def producer(queue):
    for i in range(3):
        await queue.put(i)

async def consumer(queue, out):
    while True:
        item = await queue.get()
        out.append(item)
        queue.task_done()
        if len(out) == 3:
            break

async def main():
    q = asyncio.Queue()
    out = []
    prod_task = asyncio.create_task(producer(q))
    cons_task = asyncio.create_task(consumer(q, out))
    await prod_task
    await cons_task
    print(out)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 2]"]);
}

#[test]
fn test_py_async_taskgroup_py311() {
    let src = r#"
import asyncio, sys

async def worker(name, delay, out):
    await asyncio.sleep(delay)
    out.append(name)

async def main():
    out = []
    if sys.version_info >= (3, 11):
        async with asyncio.TaskGroup() as tg:
            tg.create_task(worker("w1", 0.002, out))
            tg.create_task(worker("w2", 0.001, out))
    else:
        await asyncio.gather(worker("w2", 0.001, out), worker("w1", 0.002, out))
    print("w2" in out and "w1" in out)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_async_semaphore_concurrency_limit() {
    let src = r#"
import asyncio

active = 0
max_active = 0

async def worker(sem):
    global active, max_active
    async with sem:
        active += 1
        if active > max_active:
            max_active = active
        await asyncio.sleep(0.001)
        active -= 1

async def main():
    sem = asyncio.Semaphore(2)
    tasks = [asyncio.create_task(worker(sem)) for _ in range(5)]
    await asyncio.gather(*tasks)
    print(max_active <= 2)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_async_shield_cancellation_protection() {
    let src = r#"
import asyncio

async def critical_task():
    await asyncio.sleep(0.01)
    return "saved"

async def main():
    task = asyncio.create_task(asyncio.shield(critical_task()))
    await asyncio.sleep(0.001)
    task.cancel()
    try:
        await task
    except asyncio.CancelledError:
        print("Wrapper cancelled")
    val = await critical_task()
    print(val)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["Wrapper cancelled", "saved"]);
}

#[test]
fn test_py_async_run_coroutine_threadsafe() {
    let src = r#"
import asyncio, threading

async def async_add(a, b):
    return a + b

def thread_func(loop, out):
    fut = asyncio.run_coroutine_threadsafe(async_add(10, 20), loop)
    out.append(fut.result())

async def main():
    loop = asyncio.get_running_loop()
    out = []
    t = threading.Thread(target=thread_func, args=(loop, out))
    t.start()
    t.join()
    print(out)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["[30]"]);
}

#[test]
fn test_py_async_event_loop_time_counter() {
    let src = r#"
import asyncio

async def main():
    loop = asyncio.get_running_loop()
    t1 = loop.time()
    await asyncio.sleep(0.001)
    t2 = loop.time()
    print(t2 > t1)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_async_condition_wait_notify() {
    let src = r#"
import asyncio

async def consumer(cond, state):
    async with cond:
        await cond.wait()
        state.append("notified")

async def producer(cond):
    await asyncio.sleep(0.001)
    async with cond:
        cond.notify()

async def main():
    cond = asyncio.Condition()
    state = []
    t_cons = asyncio.create_task(consumer(cond, state))
    t_prod = asyncio.create_task(producer(cond))
    await t_cons
    await t_prod
    print(state)

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["['notified']"]);
}

#[test]
fn test_py_async_as_completed_iterator() {
    let src = r#"
import asyncio

async def delay_return(val, delay):
    await asyncio.sleep(delay)
    return val

async def main():
    tasks = [delay_return(3, 0.003), delay_return(1, 0.001), delay_return(2, 0.002)]
    results = []
    for fut in asyncio.as_completed(tasks):
        res = await fut
        results.append(res)
    print(results)  # should arrive in order of completion: [1, 2, 3]

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]"]);
}

#[test]
fn test_py_async_wait_first_completed() {
    let src = r#"
import asyncio

async def fast(): return "fast"
async def slow(): await asyncio.sleep(10); return "slow"

async def main():
    t1 = asyncio.create_task(fast())
    t2 = asyncio.create_task(slow())
    done, pending = await asyncio.wait([t1, t2], return_when=asyncio.FIRST_COMPLETED)
    print(len(done))
    print(next(iter(done)).result())
    t2.cancel()

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["1", "fast"]);
}

#[test]
fn test_py_async_current_task_inspection() {
    let src = r#"
import asyncio

async def check_task():
    task = asyncio.current_task()
    return task is not None

async def main():
    print(await check_task())

asyncio.run(main())
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
