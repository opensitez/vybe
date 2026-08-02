// vybe-test: js/ecma_operators/loose_equality_with_booleans_and_strings
// origin: languages/js/tests/js/test_ecma_operators.rs

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

__check(__line("0" == false), "true");
__check(__line("1" == true), "true");
__check(__line(2 == true), "false");
