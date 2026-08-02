// vybe-test: js/do_while_while_for_loop_control_flow/test_js_for_in_loop_null_or_undefined_is_noop
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

let executed = false;
for (const k in null) executed = true;
for (const k in undefined) executed = true;
console.log(executed);
