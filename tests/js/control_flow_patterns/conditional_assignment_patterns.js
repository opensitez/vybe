// vybe-test: js/control_flow_patterns/conditional_assignment_patterns
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

let x;
x = x || 10;      // short-circuit assignment (before logical assign)
__check(__line(x), "10");
x = x && (x * 2);
__check(__line(x), "20");
