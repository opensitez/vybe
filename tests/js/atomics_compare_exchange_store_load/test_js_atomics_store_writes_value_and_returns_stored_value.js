// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_store_writes_value_and_returns_stored_value
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
const stored = Atomics.store(i32, 0, 99);
__check(__line(stored + "|" + Atomics.load(i32, 0)), "99|99");
