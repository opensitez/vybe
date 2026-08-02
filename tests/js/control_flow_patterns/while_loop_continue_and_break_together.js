// vybe-test: js/control_flow_patterns/while_loop_continue_and_break_together
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
const values = [];
while (n < 6) {
    n++;
    if (n % 2 === 0) continue;
    if (n > 4) break;
    values.push(n);
}
console.log(values.join(","));
