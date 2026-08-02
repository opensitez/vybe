// vybe-test: js/async_try_catch_finally_rejection_propagation/test_js_async_catch_destructuring_error_object
// origin: languages/js/tests/js/test_js_async_try_catch_finally_rejection_propagation.rs

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

async function destructureError() {
    try {
        const err = new Error("BadFormat");
        err.code = 400;
        throw err;
    } catch ({ message, code }) {
        __check(__line(`${message}:${code}`), "BadFormat:400");
    }
}
destructureError();
