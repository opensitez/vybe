// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_same_value_zero_typed_array_includes
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

const floats = new Float64Array([+0, NaN]);
__check(__line(`${floats.includes(-0)}:${floats.includes(NaN)}`), "true:true");
