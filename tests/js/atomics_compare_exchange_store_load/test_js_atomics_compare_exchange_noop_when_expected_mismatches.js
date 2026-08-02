// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_compare_exchange_noop_when_expected_mismatches
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

const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 10;
const old = Atomics.compareExchange(i32, 0, 99, 20); // expected 99 != current 10 -> no swap!
__check(__line(old + "|" + i32[0]), "10|10");
