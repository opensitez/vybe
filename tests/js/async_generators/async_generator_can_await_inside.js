// vybe-test: js/async_generators/async_generator_can_await_inside
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

function delay(v) { return Promise.resolve(v * 10); }
async function* gen() {
    yield await delay(1);
    yield await delay(2);
    yield await delay(3);
}
async function main() {
    const results = [];
    for await (const v of gen()) results.push(v);
    console.log(results.join(","));
}
main();
