// vybe-test: js/atomics_advanced/atomics_wait_returns_not_equal_when_value_differs
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
view[0] = 1;
__check(__line(Atomics.wait(view, 0, 0, 1)), "not-equal");
