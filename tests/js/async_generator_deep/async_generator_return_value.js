// vybe-test: js/async_generator_deep/async_generator_return_value
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
    return "done";
}
async function main() {
    const it = gen();
    const r1 = await it.next();
    const r2 = await it.next();
    console.log(r1.value);
    console.log(r2.value);
    console.log(r2.done);
}
main();
