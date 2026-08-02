// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_map_returns_same_typedarray_constructor
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

const u8 = new Uint8Array([1, 2, 3]);
const mapped = u8.map(x => x * 10);
__check(__line(mapped.join(",") + "|isUint8=" + (mapped instanceof Uint8Array)), "10,20,30|isUint8=true");
