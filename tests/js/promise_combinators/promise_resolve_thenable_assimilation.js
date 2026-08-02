// vybe-test: js/promise_combinators/promise_resolve_thenable_assimilation
// origin: languages/js/tests/js/test_promise_combinators.rs

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

async function main() {
    const thenable = { then(resolve) { resolve(100); } };
    const result = await Promise.resolve(thenable);
    console.log(result);
}
main();
