// vybe-test: js/typed_array_from_of_constructors/test_js_typedarray_from_typedarray_reencoding
// origin: languages/js/tests/js/test_js_typed_array_from_of_constructors.rs

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

const f32 = new Float32Array([1.5, 2.5]);
const u8 = Uint8Array.from(f32);
__check(__line(u8.join(",")), "1,2");
