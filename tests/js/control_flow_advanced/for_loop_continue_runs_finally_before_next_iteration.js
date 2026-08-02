// vybe-test: js/control_flow_advanced/for_loop_continue_runs_finally_before_next_iteration
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
for (let i = 0; i < 3; i++) {
    try {
        if (i === 1) {
            continue;
        }
        log.push("body-" + i);
    } finally {
        log.push("finally-" + i);
    }
}
console.log(log.join(","));
