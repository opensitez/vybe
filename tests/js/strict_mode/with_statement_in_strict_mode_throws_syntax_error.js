// vybe-test: js/strict_mode/with_statement_in_strict_mode_throws_syntax_error
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
    eval('"use strict"; with ({}) {}');
} catch (e) {
    threw = true;
}
__check(__line(threw), "true");
