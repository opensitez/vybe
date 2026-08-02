// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_constructor_copy_from_another_typedarray
// origin: languages/js/tests/js/test_js_typed_array_uint8_int32_float64_views.rs

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

const original = new Float32Array([1.5, 2.5]);
const copy = new Float64Array(original);
__check(__line(copy.length + "|" + copy[0] + "|" + (copy.buffer !== original.buffer)), "2|1.5|true");
