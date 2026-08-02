// vybe-test: js/async_concurrency_patterns/deferred_promise
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
