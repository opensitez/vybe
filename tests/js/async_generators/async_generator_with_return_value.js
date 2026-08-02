// vybe-test: js/async_generators/async_generator_with_return_value
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
    yield "a";
    return "final";
}
async function main() {
    const g = gen();
    const r1 = await g.next();
    const r2 = await g.next();
    const r3 = await g.next();
    console.log(r1.value + "," + r1.done);
    console.log(r2.value + "," + r2.done);
    console.log(r3.value + "," + r3.done);
}
main();
