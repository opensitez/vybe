// vybe-test: js/do_while_while_for_loop_control_flow/test_js_continue_statement_skips_iteration
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

const log = [];
for (let i = 0; i < 5; i++) {
    if (i % 2 === 0) continue;
    log.push(i);
}
console.log(log.join(","));
