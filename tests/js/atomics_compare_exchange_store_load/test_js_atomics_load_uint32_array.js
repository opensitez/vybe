// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_load_uint32_array
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

const u32 = new Uint32Array(new SharedArrayBuffer(4));
Atomics.store(u32, 0, 4294967295);
__check(__line(Atomics.load(u32, 0)), "4294967295");
