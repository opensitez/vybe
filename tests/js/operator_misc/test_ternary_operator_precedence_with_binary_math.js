// vybe-test: js/operator_misc/test_ternary_operator_precedence_with_binary_math
// origin: languages/js/tests/js/test_operator_misc.rs

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

__check(__line((1 + 2 ? 3 : 4 * 5) + "|" + (0 ? 3 : 4 * 5)), "3|20");
