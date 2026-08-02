// vybe-test: js/async_generators/async_generator_wraps_sync_generator
// origin: languages/js/tests/js/test_async_generators.rs

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

function* syncGen() {
    yield 1; yield 2; yield 3;
}
async function* asyncWrap(iter) {
    for (const v of iter) yield await Promise.resolve(v * 10);
}
async function main() {
    const results = [];
    for await (const v of asyncWrap(syncGen())) results.push(v);
    console.log(results.join(","));
}
main();
