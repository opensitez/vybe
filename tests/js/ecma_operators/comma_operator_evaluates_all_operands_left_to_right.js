// vybe-test: js/ecma_operators/comma_operator_evaluates_all_operands_left_to_right
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

const seen = [];
const value = (seen.push("a"), seen.push("b"), 5);
__check(__line(value), "5");
__check(__line(seen.join(",")), "a,b");
