// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_biguint64array_deep_copy
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

const bu64 = new BigUint64Array([0xFFFFFFFFFFFFFFFFn]);
const clone = structuredClone(bu64);
__check(__line((clone instanceof BigUint64Array) + "|" + clone[0].toString(16)), "true|ffffffffffffffff");
