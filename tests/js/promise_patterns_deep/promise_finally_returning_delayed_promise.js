// vybe-test: js/promise_patterns_deep/promise_finally_returning_delayed_promise
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
    const res = await Promise.resolve("done").finally(() => new Promise(r => setTimeout(r, 10)));
    console.log(res);
}
main();
