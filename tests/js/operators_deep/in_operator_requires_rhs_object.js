// vybe-test: js/operators_deep/in_operator_requires_rhs_object
// origin: languages/js/tests/js/test_operators_deep.rs

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

__check(__line("x" in { x: 1, y: 2 }), "true");
__check(__line("x" in 1), "false");
__check(__line("x" in Object(1)), "false");
