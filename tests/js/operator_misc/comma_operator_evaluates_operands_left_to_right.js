// vybe-test: js/operator_misc/comma_operator_evaluates_operands_left_to_right
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

const log = [];
const value = (log.push("a"), log.push("b"), log.push("c"), "final");
__check(__line(value), "final");
__check(__line(log.join(",")), "1,2,3");
