// vybe-test: js/async_generator_deep/async_generator_early_return
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
    try {
        yield 1;
        yield 2;
        yield 3;
    } finally {
        console.log("cleanup");
    }
}
async function main() {
    const it = gen();
    await it.next();
    await it.return("stop");
    console.log("done");
}
main();
