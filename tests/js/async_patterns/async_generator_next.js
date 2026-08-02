// vybe-test: js/async_patterns/async_generator_next
// origin: languages/js/tests/js/test_async_patterns.rs

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
    yield 10;
    yield 20;
}
async function main() {
    let g = gen();
    let r1 = await g.next();
    console.log(r1.value);
    console.log(r1.done);
    let r2 = await g.next();
    console.log(r2.value);
    let r3 = await g.next();
    console.log(r3.done);
}
main();
