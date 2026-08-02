// vybe-test: js/control_flow_patterns/while_loop_condition_increments_and_stops
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

let n = 0;
let iterations = 0;
while ((n += 1) < 4) {
    iterations++;
}
console.log(iterations);
console.log(n);
