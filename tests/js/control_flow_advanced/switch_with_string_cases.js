// vybe-test: js/control_flow_advanced/switch_with_string_cases
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

function grade(score) {
    switch (true) {
        case score >= 90: return "A";
        case score >= 80: return "B";
        case score >= 70: return "C";
        default: return "F";
    }
}
__check(__line(grade(95)), "A");
__check(__line(grade(82)), "B");
__check(__line(grade(60)), "F");
