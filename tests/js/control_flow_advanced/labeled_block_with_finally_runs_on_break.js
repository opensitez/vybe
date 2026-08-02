// vybe-test: js/control_flow_advanced/labeled_block_with_finally_runs_on_break
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

let steps = [];
outer: {
    try {
        steps.push("before");
        break outer;
    } finally {
        steps.push("finally");
    }
}
__check(__line(steps.join("|")), "before|finally");
