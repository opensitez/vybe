// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_float32array_nan_values
// origin: languages/js/tests/js/test_js_structured_clone_typed_arrays_array_buffers.rs

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

const f32 = new Float32Array([NaN, Infinity, -Infinity]);
const clone = structuredClone(f32);
__check(__line(Number.isNaN(clone[0]) + "|" + clone[1] + "|" + clone[2]), "true|Infinity|-Infinity");
