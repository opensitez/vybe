// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_typedarray_prototype_methods_intact
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

const u8 = new Uint8Array([5, 10, 15]);
const clone = structuredClone(u8);
const mapped = clone.map(x => x * 2);
__check(__line(mapped.join(",")), "10,20,30");
