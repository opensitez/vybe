// vybe-test: js/operators_deep/exponentiation_precedence_and_syntax_boundary
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

__check(__line((2 ** 3) ** 2), "64"); // explicit grouping
__check(__line(2 ** 3 ** 2), "64");    // right-associative exponentiation
__check(__line("end"), "end");
