// vybe-test: js/do_while_while_for_loop_control_flow/test_js_while_loop_condition_side_effects_and_continue
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
let checks = 0;
while ((checks++, i < 3)) {
    if (i === 1) {
        i++;
        continue;
    }
    i++;
}
console.log(i + "|" + checks);
