// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_object_is_primitive_comparisons
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

__check(__line(`${Object.is("a", "a")}:${Object.is(true, true)}:${Object.is(null, null)}:${Object.is(undefined, undefined)}`), "true:true:true:true");
