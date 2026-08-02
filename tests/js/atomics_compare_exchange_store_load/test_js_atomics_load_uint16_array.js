// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_load_uint16_array
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

const u16 = new Uint16Array(new SharedArrayBuffer(2));
Atomics.store(u16, 0, 65535);
__check(__line(Atomics.load(u16, 0)), "65535");
