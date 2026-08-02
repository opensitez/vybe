// vybe-test: js/control_flow_advanced/for_loop_finally_runs_before_break
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
        events.push("loop-" + i);
        if (i === 1) {
            break;
        }
    } finally {
        events.push("finally-" + i);
    }
}
console.log(events.join(","));
