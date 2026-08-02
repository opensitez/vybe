// vybe-test: js/do_while_while_for_loop_control_flow/test_js_do_while_with_continue_in_finally_scope
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
        if (i === 0) {
            i++;
            continue;
        }
        if (i === 1) {
            break;
        }
        i++;
    } finally {
        log.push("finally" + i);
    }
} while (i < 5);
console.log(log.join("|"));
