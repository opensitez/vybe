// vybe-test: js/function_declaration_hoisting_in_blocks/test_js_function_declaration_in_try_block_strict_mode
// origin: languages/js/tests/js/test_js_function_declaration_hoisting_in_blocks.rs

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

"use strict";
try {
    function tryFunc() { return "TryFunc"; }
    __check(__line(tryFunc()), "TryFunc");
} catch (e) {}
__check(__line(typeof tryFunc), "undefined");
