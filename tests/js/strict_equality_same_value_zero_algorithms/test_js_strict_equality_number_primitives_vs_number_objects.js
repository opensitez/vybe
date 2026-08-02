// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_strict_equality_number_primitives_vs_number_objects
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

const numPrim = 42;
const numObj = new Number(42);
__check(__line(`${numPrim === 42}:${numPrim === numObj}`), "true:false");
