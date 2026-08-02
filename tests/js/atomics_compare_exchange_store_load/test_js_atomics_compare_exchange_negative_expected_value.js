// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_compare_exchange_negative_expected_value
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

const i8 = new Int8Array(new SharedArrayBuffer(1));
i8[0] = -50;
const old = Atomics.compareExchange(i8, 0, -50, 100);
__check(__line(old + "|" + i8[0]), "-50|100");
