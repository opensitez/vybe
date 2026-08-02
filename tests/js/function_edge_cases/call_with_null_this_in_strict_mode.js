// vybe-test: js/function_edge_cases/call_with_null_this_in_strict_mode
// origin: languages/js/tests/js/test_function_edge_cases.rs

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

function whoami() {
    "use strict";
    return this === null ? "null-this" : "other";
}
__check(__line(whoami.call(null)), "null-this");
