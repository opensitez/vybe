// vybe-test: js/async_generators/async_generator_return_method_finishes_early
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
    yield 2;
    yield 3;
}
async function main() {
    const g = gen();
    const r1 = await g.next();
    const ret = await g.return(99);
    const r2 = await g.next();
    console.log(r1.value);
    console.log(ret.value);
    console.log(r2.done);
}
main();
