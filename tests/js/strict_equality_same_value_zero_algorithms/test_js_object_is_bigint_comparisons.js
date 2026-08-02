// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_object_is_bigint_comparisons
// origin: languages/js/tests/js/test_js_strict_equality_same_value_zero_algorithms.rs

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

__check(__line(`${Object.is(100n, 100n)}:${Object.is(100n, 100)}:${Object.is(0n, -0n)}`), "true:false:true");
