// vybe-test: js/control_flow_advanced/switch_default_at_end_runs_when_no_match
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

let x = "z";
switch (x) {
    case "a": console.log("a"); break;
    case "b": console.log("b"); break;
    default: console.log("other");
}
