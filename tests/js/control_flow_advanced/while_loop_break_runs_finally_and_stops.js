// vybe-test: js/control_flow_advanced/while_loop_break_runs_finally_and_stops
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
        if (i === 1) {
            break;
        }
        log.push("body-" + i);
        i++;
    } finally {
        log.push("finally-" + i);
    }
}
console.log(log.join("|"));
console.log(i);
