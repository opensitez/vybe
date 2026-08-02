// vybe-test: js/async_generators/async_generator_throw_before_first_next
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
    yield 1;
}
async function main() {
    const g = gen();
    try {
        await g.throw(new Error("early_throw"));
    } catch (e) {
        console.log(e.message);
    }
}
main();
