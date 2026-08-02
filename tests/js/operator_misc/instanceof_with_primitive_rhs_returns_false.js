// vybe-test: js/operator_misc/instanceof_with_primitive_rhs_returns_false
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

let result = "unreached";
try {
    result = ({} instanceof 7) ? "true" : "false";
} catch (e) {
    result = "error";
}
__check(__line(result), "false");
