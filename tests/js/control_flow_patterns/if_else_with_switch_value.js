// vybe-test: js/control_flow_patterns/if_else_with_switch_value
// origin: languages/js/tests/js/test_control_flow_patterns.rs

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

const n = 1;
let out;
if (n === 0) {
    out = "zero";
} else if (n === 1) {
    out = "one";
} else {
    out = "other";
}
__check(__line(out), "one");
