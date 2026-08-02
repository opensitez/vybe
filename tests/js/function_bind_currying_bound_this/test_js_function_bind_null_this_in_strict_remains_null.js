// vybe-test: js/function_bind_currying_bound_this/test_js_function_bind_null_this_in_strict_remains_null
// origin: languages/js/tests/js/test_js_function_bind_currying_bound_this.rs

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

function getThisNull() {
    "use strict";
    return this === null;
}
const boundNull = getThisNull.bind(null);
__check(__line(boundNull()), "true");
