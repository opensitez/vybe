// vybe-test: js/atomics_is_lock_free_buffer_size/test_js_atomics_is_lock_free_infinity_returns_false
// origin: languages/js/tests/js/test_js_atomics_is_lock_free_buffer_size.rs

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

__check(__line(Atomics.isLockFree(Infinity)), "false");
