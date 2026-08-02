// vybe-test: js/control_flow_patterns/switch_default_at_middle
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

function test(x) {
    switch (x) {
        case 1: return "one";
        default: return "other";
        case 2: return "two"; // still reachable
    }
}
__check(__line(test(1)), "one");
__check(__line(test(2)), "two");
__check(__line(test(99)), "other");
