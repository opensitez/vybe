// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_store_out_of_bounds_throws_rangeerror
// origin: languages/js/tests/js/test_js_atomics_compare_exchange_store_load.rs

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

const i32 = new Int32Array(new SharedArrayBuffer(4));
try {
    Atomics.store(i32, -1, 10);
} catch (e) {
    __check(__line("Atomics.store Out of Bounds RangeError"), "Atomics.store Out of Bounds RangeError");
}
