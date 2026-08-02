// vybe-test: js/control_flow_patterns/if_condition_uses_comma_expression_evaluation_order
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

const trace = [];
const mark = (name, value) => {
    trace.push(name);
    return value;
};

let x = 0;
if (mark("lhs", x = 5), mark("rhs", x > 2)) {
    trace.push("then");
}

console.log(trace.join(","));
console.log(x);
