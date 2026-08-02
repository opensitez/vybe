// vybe-test: js/function_bind_currying_bound_this/test_js_function_bind_primitive_this_unboxed_in_strict_mode
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

function checkThisType() {
    "use strict";
    return typeof this;
}
const boundNum = checkThisType.bind(42);
__check(__line(boundNum()), "number"); // In strict mode, primitive this remains primitive number!
