// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_int32array_deep_copy
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

const i32 = new Int32Array([100, -200, 300]);
const clone = structuredClone(i32);
__check(__line((clone instanceof Int32Array) + "|" + clone.join(",")), "true|100,-200,300");
