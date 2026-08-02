// vybe-test: js/number_edge_basics/number_is_integer_and_safe_integer_checks
// origin: languages/js/tests/js/test_number_edge_basics.rs

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

__check(__line(Number.isInteger(10)), "true");
__check(__line(Number.isInteger(10.1)), "false");
__check(__line(Number.isSafeInteger(9007199254740991)), "true");
__check(__line(Number.isSafeInteger(9007199254740992)), "false");
