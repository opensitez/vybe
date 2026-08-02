// vybe-test: js/async_concurrency_patterns/async_iterator_custom
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
