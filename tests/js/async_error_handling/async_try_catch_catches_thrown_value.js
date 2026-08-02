// vybe-test: js/async_error_handling/async_try_catch_catches_thrown_value
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
    try {
        throw new Error("async error");
    } catch (e) {
        return e.message;
    }
}
f().then(v => console.log(v));
