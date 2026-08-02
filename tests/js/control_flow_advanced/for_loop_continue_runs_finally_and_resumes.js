// vybe-test: js/control_flow_advanced/for_loop_continue_runs_finally_and_resumes
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

let events = [];
for (let i = 0; i < 3; i++) {
    try {
        if (i === 1) {
            continue;
        }
        events.push("work-" + i);
    } finally {
        events.push("finally-" + i);
    }
}
console.log(events.join(","));
