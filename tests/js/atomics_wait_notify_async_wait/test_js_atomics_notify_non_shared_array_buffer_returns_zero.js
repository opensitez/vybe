// vybe-test: js/atomics_wait_notify_async_wait/test_js_atomics_notify_non_shared_array_buffer_returns_zero
// origin: languages/js/tests/js/test_js_atomics_wait_notify_async_wait.rs

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

const i32 = new Int32Array(1);
__check(__line(Atomics.notify(i32, 0, 1)), "0");
