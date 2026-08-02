// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_relational_evaluation_order_left_to_right_side_effects
// origin: languages/js/tests/js/test_js_relational_in_instanceof_less_greater_operators.rs

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
const left = {
    [Symbol.toPrimitive]() {
        log.push("L");
        return 10;
    }
};
const right = {
    [Symbol.toPrimitive]() {
        log.push("R");
        return 5;
    }
};
__check(__line((left > right) + "|" + log.join(",")), "true|L,R");
