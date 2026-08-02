// vybe-test: js/control_flow_advanced/switch_strict_equality_no_coercion
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

let val = "1";
switch (val) {
    case 1: console.log("number"); break;
    case "1": console.log("string"); break;
}
