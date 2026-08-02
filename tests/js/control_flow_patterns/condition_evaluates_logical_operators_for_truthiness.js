// vybe-test: js/control_flow_patterns/condition_evaluates_logical_operators_for_truthiness
// origin: languages/js/tests/js/test_control_flow_patterns.rs

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

const calls = [];
const side_effect = () => {
    calls.push("side");
    return true;
};

if (0 && side_effect()) {
    calls.push("if-true");
} else {
    calls.push("if-false");
}

if (1 || side_effect()) {
    calls.push("or-true");
}

__check(__line(calls.join(",")), "if-false,or-true");
