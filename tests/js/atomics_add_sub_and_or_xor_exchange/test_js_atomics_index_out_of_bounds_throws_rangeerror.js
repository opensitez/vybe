// vybe-test: js/atomics_add_sub_and_or_xor_exchange/test_js_atomics_index_out_of_bounds_throws_rangeerror
// origin: languages/js/tests/js/test_js_atomics_add_sub_and_or_xor_exchange.rs

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

const i32 = new Int32Array(2);
try {
    Atomics.add(i32, 5, 1);
} catch (e) {
    __check(__line("Atomics Index Out of Bounds RangeError"), "Atomics Index Out of Bounds RangeError");
}
