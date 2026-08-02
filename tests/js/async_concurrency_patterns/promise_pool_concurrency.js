// vybe-test: js/async_concurrency_patterns/promise_pool_concurrency
// origin: languages/js/tests/js/test_async_concurrency_patterns.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

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
