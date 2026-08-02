// vybe-test: js/async_generators/async_generator_next_accepts_value
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

async function* gen() {
    const x = yield "first";
    yield "second:" + x;
}
async function main() {
    const g = gen();
    await g.next();
    const r = await g.next("hello");
    console.log(r.value);
}
main();
