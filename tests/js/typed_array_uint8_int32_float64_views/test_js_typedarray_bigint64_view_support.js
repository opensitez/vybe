// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_bigint64_view_support
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

const big64 = new BigInt64Array(2);
big64[0] = 9007199254740991n;
__check(__line(big64[0].toString()), "9007199254740991");
