// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_biguint64_view_overflow
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

const bigu64 = new BigUint64Array(1);
bigu64[0] = 0xFFFFFFFFFFFFFFFFn;
__check(__line(bigu64[0].toString()), "18446744073709551615");
