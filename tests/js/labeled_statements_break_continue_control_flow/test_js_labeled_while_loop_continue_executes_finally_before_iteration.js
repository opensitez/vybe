// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_labeled_while_loop_continue_executes_finally_before_iteration
// origin: languages/js/tests/js/test_js_labeled_statements_break_continue_control_flow.rs

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
const log = [];

outer: while (i < 4) {
    try {
        log.push("body-" + i);
        if (i === 1) {
            i += 1;
            continue outer;
        }
        log.push("work-" + i);
        i++;
    } finally {
        log.push("finally-" + i);
    }
}

console.log(log.join("|"));
