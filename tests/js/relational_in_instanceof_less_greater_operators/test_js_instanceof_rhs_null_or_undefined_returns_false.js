// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_instanceof_rhs_null_or_undefined_returns_false
// origin: languages/js/tests/js/test_js_relational_in_instanceof_less_greater_operators.rs

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

const obj = {};
try {
    console.log(obj instanceof null);
} catch (e) {
    console.log("null");
}
try {
    console.log(obj instanceof undefined);
} catch (e) {
    console.log("undefined");
}
