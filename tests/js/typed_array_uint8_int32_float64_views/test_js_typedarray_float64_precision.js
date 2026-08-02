// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_float64_precision
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

const f64 = new Float64Array(2);
f64[0] = 3.141592653589793;
f64[1] = NaN;
__check(__line(f64[0] + "|" + Number.isNaN(f64[1])), "3.141592653589793|true");
