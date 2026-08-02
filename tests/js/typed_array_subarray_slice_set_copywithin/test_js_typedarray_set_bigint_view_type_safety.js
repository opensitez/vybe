// vybe-test: js/typed_array_subarray_slice_set_copywithin/test_js_typedarray_set_bigint_view_type_safety
// origin: languages/js/tests/js/test_js_typed_array_subarray_slice_set_copywithin.rs

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

const big = new BigInt64Array(2);
big.set([100n, 200n]);
__check(__line(big[0].toString() + "|" + big[1].toString()), "100|200");
