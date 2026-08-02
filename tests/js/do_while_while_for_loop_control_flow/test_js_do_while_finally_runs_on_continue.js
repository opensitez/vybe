// vybe-test: js/do_while_while_for_loop_control_flow/test_js_do_while_finally_runs_on_continue
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

let log = [];
let i = 0;
do {
    try {
        log.push("body" + i);
        if (i === 1) {
            i++;
            continue;
        }
        i++;
    } finally {
        log.push("finally" + i);
    }
} while (i < 4);
console.log(log.join("|"));
