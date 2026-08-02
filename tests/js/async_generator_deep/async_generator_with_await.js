// vybe-test: js/async_generator_deep/async_generator_with_await
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
    const a = await Promise.resolve(10);
    yield a;
    const b = await Promise.resolve(20);
    yield b;
}
async function main() {
    const results = [];
    for await (const v of gen()) results.push(v);
    console.log(results.join(","));
}
main();
