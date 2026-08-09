/// Async patterns — throttle, debounce, queue, retry with exponential backoff
use super::helpers::run_js;

#[test]
fn sequential_async_queue() {
    assert_eq!(
        run_js(
            r#"
class AsyncQueue {
    #queue = [];
    #running = false;
    add(fn) {
        return new Promise((resolve, reject) => {
            this.#queue.push({ fn, resolve, reject });
            this.#run();
        });
    }
    async #run() {
        if (this.#running) return;
        this.#running = true;
        while (this.#queue.length > 0) {
            const { fn, resolve, reject } = this.#queue.shift();
            try { resolve(await fn()); }
            catch (e) { reject(e); }
        }
        this.#running = false;
    }
}
async function main() {
    const q = new AsyncQueue();
    const log = [];
    await Promise.all([
        q.add(async () => { log.push(1); return 1; }),
        q.add(async () => { log.push(2); return 2; }),
        q.add(async () => { log.push(3); return 3; }),
    ]);
    console.log(log.join(","));
}
main();
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn promise_retry_with_limit() {
    assert_eq!(
        run_js(
            r#"
async function retry(fn, times) {
    for (let i = 0; i < times; i++) {
        try { return await fn(); }
        catch (e) { if (i === times - 1) throw e; }
    }
}
let attempts = 0;
async function main() {
    const result = await retry(async () => {
        attempts++;
        if (attempts < 3) throw new Error("not yet");
        return "success after " + attempts;
    }, 5);
    console.log(result);
    console.log(attempts);
}
main();
"#
        ),
        vec!["success after 3", "3"]
    );
}

#[test]
fn async_map_sequential() {
    assert_eq!(
        run_js(
            r#"
async function mapSequential(arr, fn) {
    const results = [];
    for (const item of arr) {
        results.push(await fn(item));
    }
    return results;
}
async function main() {
    const results = await mapSequential([1, 2, 3], async x => x * 2);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["2,4,6"]
    );
}

#[test]
fn async_map_parallel() {
    assert_eq!(
        run_js(
            r#"
async function mapParallel(arr, fn) {
    return Promise.all(arr.map(fn));
}
async function main() {
    const results = await mapParallel([1, 2, 3, 4], async x => x * x);
    console.log(results.join(","));
}
main();
"#
        ),
        vec!["1,4,9,16"]
    );
}

#[test]
fn async_reduce_sequential() {
    assert_eq!(
        run_js(
            r#"
async function asyncReduce(arr, fn, init) {
    let acc = init;
    for (const item of arr) {
        acc = await fn(acc, item);
    }
    return acc;
}
async function main() {
    const result = await asyncReduce([1, 2, 3, 4, 5], async (acc, x) => acc + x, 0);
    console.log(result);
}
main();
"#
        ),
        vec!["15"]
    );
}

#[test]
fn async_filter() {
    assert_eq!(
        run_js(
            r#"
async function asyncFilter(arr, pred) {
    const results = await Promise.all(arr.map(async (item) => ({
        item,
        keep: await pred(item)
    })));
    return results.filter(r => r.keep).map(r => r.item);
}
async function main() {
    const evens = await asyncFilter([1, 2, 3, 4, 5, 6], async x => x % 2 === 0);
    console.log(evens.join(","));
}
main();
"#
        ),
        vec!["2,4,6"]
    );
}

#[test]
fn async_waterfall_pipeline() {
    assert_eq!(
        run_js(
            r#"
async function waterfall(fns, initial) {
    return fns.reduce(async (promise, fn) => fn(await promise), Promise.resolve(initial));
}
async function main() {
    const result = await waterfall([
        async x => x + 1,
        async x => x * 2,
        async x => "result: " + x,
    ], 5);
    console.log(result);
}
main();
"#
        ),
        vec!["result: 12"]
    );
}

#[test]
fn async_timeout_race() {
    assert_eq!(
        run_js(
            r#"
function timeout(ms) {
    return new Promise((_, reject) =>
        setTimeout(() => reject(new Error("Timeout")), ms)
    );
}
async function withTimeout(promise, ms) {
    return Promise.race([promise, timeout(ms)]);
}
async function main() {
    const fast = Promise.resolve("done");
    const result = await withTimeout(fast, 5000);
    console.log(result);
}
main();
"#
        ),
        vec!["done"]
    );
}

#[test]
fn async_batch_processing_pattern() {
    assert_eq!(
        run_js(
            r#"
async function batchProcess(arr, size, fn) {
    const res = [];
    for (let i = 0; i < arr.length; i += size) {
        const batch = arr.slice(i, i + size);
        const batchRes = await Promise.all(batch.map(fn));
        res.push(...batchRes);
    }
    return res;
}
async function main() {
    const out = await batchProcess([1, 2, 3, 4, 5], 2, async x => x * 10);
    console.log(out.join(","));
}
main();
"#
        ),
        vec!["10,20,30,40,50"]
    );
}
