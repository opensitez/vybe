// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_store_coerces_value_to_typed_array_type
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

const u8 = new Uint8Array(new SharedArrayBuffer(1));
Atomics.store(u8, 0, "150");
__check(__line(Atomics.load(u8, 0)), "150");
