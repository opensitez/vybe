// vybe-test: js/do_while_while_for_loop_control_flow/test_js_while_loop_zero_executions_when_condition_false
// origin: languages/js/tests/js/test_js_do_while_while_for_loop_control_flow.rs

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

let executed = 0;
while (false) {
    executed++;
}
console.log(executed);
