// vybe-test: js/control_flow_advanced/while_loop_continue_and_break_control_flow
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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

let i = 0;
const values = [];
while (i < 6) {
    i++;
    if (i === 2) continue;
    if (i === 5) break;
    values.push(i);
}
console.log(values.join(","));
