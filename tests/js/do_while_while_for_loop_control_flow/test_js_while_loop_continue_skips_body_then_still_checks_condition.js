// vybe-test: js/do_while_while_for_loop_control_flow/test_js_while_loop_continue_skips_body_then_still_checks_condition
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

const values = [];
let i = 0;
while (i < 5) {
    i++;
    if (i === 2) {
        continue;
    }
    if (i === 4) {
        break;
    }
    values.push(i);
}
console.log(values.join(","));
