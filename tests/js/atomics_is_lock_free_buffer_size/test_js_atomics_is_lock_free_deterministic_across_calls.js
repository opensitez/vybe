// vybe-test: js/atomics_is_lock_free_buffer_size/test_js_atomics_is_lock_free_deterministic_across_calls
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

const r1 = Atomics.isLockFree(4);
const r2 = Atomics.isLockFree(4);
__check(__line(r1 === r2), "true");
