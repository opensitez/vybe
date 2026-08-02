// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_uint8array_deep_copy
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

const u8 = new Uint8Array([1, 2, 3]);
const cloneU8 = structuredClone(u8);
cloneU8[0] = 99;
__check(__line(u8[0] + "|" + cloneU8[0]), "1|99");
