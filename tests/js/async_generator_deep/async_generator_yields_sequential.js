// vybe-test: js/async_generator_deep/async_generator_yields_sequential
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

const order = [];
async function* gen() {
    order.push("before 1");
    yield 1;
    order.push("before 2");
    yield 2;
}
async function main() {
    const it = gen();
    await it.next();
    order.push("after first next");
    await it.next();
    console.log(order.join(","));
}
main();
