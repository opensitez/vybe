// vybe-test: js/async_generator_deep/async_generator_yields_promises
// origin: languages/js/tests/js/test_async_generator_deep.rs

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

async function* gen() {
    yield 1;
    yield 2;
    yield 3;
}
async function main() {
    const results = [];
    for await (const v of gen()) results.push(v);
    console.log(results.join(","));
}
main();
