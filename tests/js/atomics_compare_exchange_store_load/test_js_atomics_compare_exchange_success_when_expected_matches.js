// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_compare_exchange_success_when_expected_matches
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
const old = Atomics.compareExchange(i32, 0, 10, 20); // expected 10 == current 10 -> swaps to 20!
__check(__line(old + "|" + i32[0]), "10|20");
