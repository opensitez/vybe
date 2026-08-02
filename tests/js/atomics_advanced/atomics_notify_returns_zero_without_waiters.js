// vybe-test: js/atomics_advanced/atomics_notify_returns_zero_without_waiters
// origin: languages/js/tests/js/test_atomics_advanced.rs

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

const buffer = new SharedArrayBuffer(4);
const view = new Int32Array(buffer);
__check(__line(Atomics.notify(view, 0)), "0");
