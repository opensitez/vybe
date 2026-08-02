// vybe-test: js/control_flow_patterns/if_condition_uses_assigned_value_for_truthiness
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

let x = 0;
const out = [];

if ((x = 7)) {
    out.push("if");
}
if ((x = 0)) {
    out.push("should-not-run");
} else {
    out.push("else");
}

__check(__line(out.join("|")), "if|else");
__check(__line(x), "0");
