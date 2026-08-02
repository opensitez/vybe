// vybe-test: js/promise_patterns_deep/await_unwraps_nested_promise
// origin: languages/js/tests/js/test_promise_patterns_deep.rs

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
    const p = Promise.resolve(Promise.resolve(Promise.resolve(42)));
    console.log(await p);
}
main();
