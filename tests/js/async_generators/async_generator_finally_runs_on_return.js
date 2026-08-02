// vybe-test: js/async_generators/async_generator_finally_runs_on_return
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

const log = [];
async function* gen() {
    try {
        yield 1;
    } finally {
        log.push("finally");
    }
}
async function main() {
    const g = gen();
    await g.next();
    await g.return();
    console.log(log.join(","));
}
main();
