// vybe-test: js/control_flow_patterns/switch_fallthrough_behavior
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
    let result = "";
    switch (x) {
        case 1:
            result += "1";
            // fallthrough
        case 2:
            result += "2";
            break;
        case 3:
            result += "3";
    }
    return result;
}
__check(__line(test(1)), "12"); // falls through to 2
__check(__line(test(2)), "2");
__check(__line(test(3)), "3");
