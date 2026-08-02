// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_compare_exchange_nan_in_int32_array
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
i32[0] = 0;
const old = Atomics.compareExchange(i32, 0, NaN, 100); // NaN coerces to expected 0!
__check(__line(old + "|" + i32[0]), "0|100");
