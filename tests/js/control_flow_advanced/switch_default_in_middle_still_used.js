// vybe-test: js/control_flow_advanced/switch_default_in_middle_still_used
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

let result = [];
switch (99) {
    case 1: result.push("one"); break;
    default: result.push("default");
    case 2: result.push("two"); break;
}
__check(__line(result.join(",")), "default,two");
