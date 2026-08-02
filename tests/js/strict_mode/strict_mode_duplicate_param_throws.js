// vybe-test: js/strict_mode/strict_mode_duplicate_param_throws
// origin: languages/js/tests/js/test_strict_mode.rs

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

let threw = false;
try {
    new Function('"use strict"; function f(a, a) {}; f(1,2)');
} catch (e) {
    threw = true;
}
// duplicate params in strict mode is a SyntaxError at parse time
// using new Function to test it at runtime
__check(__line(typeof threw === "boolean"), "true");
