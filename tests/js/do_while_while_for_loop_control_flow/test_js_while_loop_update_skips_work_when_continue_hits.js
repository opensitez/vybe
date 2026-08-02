// vybe-test: js/do_while_while_for_loop_control_flow/test_js_while_loop_update_skips_work_when_continue_hits
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

let i = 0;
const values = [];
while (i < 5) {
    i += 1;
    if (i === 2) {
        continue;
    }
    values.push(i);
}
console.log(values.join(","));
