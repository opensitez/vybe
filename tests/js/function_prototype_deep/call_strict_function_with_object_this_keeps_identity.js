// vybe-test: js/function_prototype_deep/call_strict_function_with_object_this_keeps_identity
// origin: languages/js/tests/js/test_function_prototype_deep.rs

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

"use strict"; function id() { return this; } const o = {}; __check(__line(id.call(o) === o), "true");
