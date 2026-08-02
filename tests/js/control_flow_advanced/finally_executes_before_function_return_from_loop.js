// vybe-test: js/control_flow_advanced/finally_executes_before_function_return_from_loop
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

function scan() {
    const events = [];
    for (let i = 0; i < 3; i++) {
        try {
            events.push("try-" + i);
            if (i === 1) return events.join("|");
        } finally {
            events.push("finally-" + i);
        }
    }
    return "after";
}
console.log(scan());
console.log("outer-start");
