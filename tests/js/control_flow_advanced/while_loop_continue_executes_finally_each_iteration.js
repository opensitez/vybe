// vybe-test: js/control_flow_advanced/while_loop_continue_executes_finally_each_iteration
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

let log = [];
let i = 0;

while (i < 3) {
    try {
        const next = i;
        if (i === 1) {
            log.push("continue-" + next);
            i += 1;
            continue;
        }
        log.push("body-" + next);
        i += 1;
    } finally {
        log.push("finally-" + i);
    }
}
console.log(log.join("|"));
