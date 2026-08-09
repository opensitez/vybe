/// Async patterns — promises, async/await, concurrency control
use super::helpers::run_js;

#[test]
fn async_queue_sequential() {
    assert_eq!(
        run_js(
            r#"
class AsyncQueue {
    #queue = [];
    #running = false;
    enqueue(task) {
        return new Promise((resolve, reject) => {
            this.#queue.push({ task, resolve, reject });
            this.#run();
        });
    }
    async #run() {
        if (this.#running) return;
        this.#running = true;
        while (this.#queue.length) {
            const { task, resolve, reject } = this.#queue.shift();
            try { resolve(await task()); } catch(e) { reject(e); }
        }
        this.#running = false;
    }
}
async function main() {
    const q = new AsyncQueue();
    const order = [];
    const results = await Promise.all([
        q.enqueue(async () => { order.push(1); return "a"; }),
        q.enqueue(async () => { order.push(2); return "b"; }),
        q.enqueue(async () => { order.push(3); return "c"; }),
    ]);
    console.log(results.join(","));
    console.log(order.join(","));
}
main();
"#
        ),
        vec!["a,b,c", "1,2,3"]
    );
}

#[test]
fn promise_pool_concurrency() {
    assert_eq!(
        run_js(
            r#"
async function pool(tasks, limit) {
    const results = [];
    const executing = [];
    for (const [i, task] of tasks.entries()) {
        const p = task().then(v => { results[i] = v; });
        executing.push(p);
        if (executing.length >= limit) await Promise.race(executing.map((p, j) => p.then(() => j)));
    }
    await Promise.all(executing);
    return results;
}
async function main() {
    const tasks = [1,2,3,4,5].map(n => () => Promise.resolve(n * 2));
    const results = await pool(tasks, 2);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["2,4,6,8,10"]
    );
}

#[test]
fn async_iterator_custom() {
    assert_eq!(
        run_js(
            r#"
async function collect(iter) {
    const result = [];
    for await (const v of iter) result.push(v);
    return result;
}
async function* range(start, end) {
    for (let i = start; i <= end; i++) yield i;
}
async function main() {
    const vals = await collect(range(1, 5));
    console.log(vals.join(","));
}
main();
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

#[test]
fn timeout_with_abort() {
    assert_eq!(
        run_js(
            r#"
function withTimeout(promise, ms) {
    let id;
    const timeout = new Promise((_, reject) => {
        id = setTimeout(() => reject(new Error("timeout")), ms);
    });
    return Promise.race([promise, timeout]).finally(() => clearTimeout(id));
}
async function main() {
    const fast = withTimeout(Promise.resolve("ok"), 1000);
    console.log(await fast);
    try {
        await withTimeout(new Promise(() => {}), 0);
    } catch(e) {
        console.log(e.message);
    }
}
main();
"#
        ),
        vec!["ok", "timeout"]
    );
}

#[test]
fn async_semaphore() {
    assert_eq!(
        run_js(
            r#"
class Semaphore {
    #count;
    #queue = [];
    constructor(count) { this.#count = count; }
    async acquire() {
        if (this.#count > 0) { this.#count--; return; }
        await new Promise(resolve => this.#queue.push(resolve));
    }
    release() {
        if (this.#queue.length) { this.#queue.shift()(); }
        else { this.#count++; }
    }
}
async function main() {
    const sem = new Semaphore(2);
    const log = [];
    async function task(id) {
        await sem.acquire();
        log.push("in:" + id);
        await Promise.resolve();
        log.push("out:" + id);
        sem.release();
    }
    await Promise.all([task(1), task(2), task(3)]);
    console.log(log.includes("in:1"));
    console.log(log.includes("in:2"));
    console.log(log.includes("out:3"));
}
main();
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn retry_with_backoff() {
    assert_eq!(
        run_js(
            r#"
async function retry(fn, maxAttempts, delay = 0) {
    for (let i = 0; i < maxAttempts; i++) {
        try { return await fn(); }
        catch(e) {
            if (i === maxAttempts - 1) throw e;
            await new Promise(r => setTimeout(r, delay));
        }
    }
}
async function main() {
    let attempts = 0;
    const result = await retry(async () => {
        attempts++;
        if (attempts < 3) throw new Error("fail");
        return "success";
    }, 5);
    console.log(result);
    console.log(attempts);
}
main();
"#
        ),
        vec!["success", "3"]
    );
}

#[test]
fn deferred_promise() {
    assert_eq!(
        run_js(
            r#"
function deferred() {
    let resolve, reject;
    const promise = new Promise((res, rej) => { resolve = res; reject = rej; });
    return { promise, resolve, reject };
}
async function main() {
    const d = deferred();
    setTimeout(() => d.resolve(42), 0);
    console.log(await d.promise);
    const d2 = deferred();
    setTimeout(() => d2.reject(new Error("fail")), 0);
    try { await d2.promise; } catch(e) { console.log(e.message); }
}
main();
"#
        ),
        vec!["42", "fail"]
    );
}

#[test]
fn async_compose() {
    assert_eq!(
        run_js(
            r#"
const asyncPipe = (...fns) => x => fns.reduce(async (p, f) => f(await p), Promise.resolve(x));
async function main() {
    const process = asyncPipe(
        async x => x + 1,
        async x => x * 2,
        async x => x.toString()
    );
    console.log(await process(5));
    console.log(await process(10));
}
main();
"#
        ),
        vec!["12", "22"]
    );
}

#[test]
fn promise_with_resolvers_static_method() {
    assert_eq!(
        run_js(
            r#"
const { promise, resolve, reject } = Promise.withResolvers();
resolve(99);
promise.then(v => console.log(v));
"#
        ),
        vec!["99"]
    );
}
