// vybe-test: js/atomics_add_sub_and_or_xor_exchange/test_js_atomics_operations_on_float_array_throws_typeerror
// origin: languages/js/tests/js/test_js_atomics_add_sub_and_or_xor_exchange.rs

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
    Atomics.add(f32, 0, 1);
} catch (e) {
    __check(__line("Atomics Float TypedArray TypeError"), "Atomics Float TypedArray TypeError");
}
