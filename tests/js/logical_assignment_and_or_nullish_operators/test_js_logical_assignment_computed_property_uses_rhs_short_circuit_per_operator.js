// vybe-test: js/logical_assignment_and_or_nullish_operators/test_js_logical_assignment_computed_property_uses_rhs_short_circuit_per_operator
// origin: languages/js/tests/js/test_js_logical_assignment_and_or_nullish_operators.rs

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

let keyEval = 0;
const key = () => {
    keyEval++;
    return "value";
};

const obj = {
    value: null,
};

obj[key()] ||= "filled"; // assigns
obj[key()] ||= "ignored"; // short-circuit, no assign
obj[key()] &&= "final";  // assigns

__check(__line(obj.value), "final");
__check(__line(keyEval), "6");
