// vybe-test: js/logical_assignment_and_or_nullish_operators/test_js_logical_assignment_const_short_circuited_still_throws_syntaxerror_or_typeerror
// origin: languages/js/tests/js/test_js_logical_assignment_and_or_nullish_operators.rs

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

const x = 5;
try {
    eval("x ||= 10;");
} catch (e) {
    __check(__line("Const Logical Assignment Reassignment Error"), "Const Logical Assignment Reassignment Error");
}
