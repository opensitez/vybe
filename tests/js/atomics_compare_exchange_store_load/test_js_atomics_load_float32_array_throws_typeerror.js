// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_load_float32_array_throws_typeerror
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

const f32 = new Float32Array(1);
try {
    Atomics.load(f32, 0);
} catch (e) {
    __check(__line("Atomics.load Float32 TypeError"), "Atomics.load Float32 TypeError");
}
