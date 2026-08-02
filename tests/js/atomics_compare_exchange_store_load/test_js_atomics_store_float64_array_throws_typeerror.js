// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_store_float64_array_throws_typeerror
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

const f64 = new Float64Array(1);
try {
    Atomics.store(f64, 0, 1.5);
} catch (e) {
    __check(__line("Atomics.store Float64 TypeError"), "Atomics.store Float64 TypeError");
}
