// vybe-test: js/async_error_handling/async_finally_does_not_change_resolved_value
// origin: languages/js/tests/js/test_async_error_handling.rs

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

async function f() {
    try { return 42; }
    finally { /* no return */ }
}
f().then(v => console.log(v));
