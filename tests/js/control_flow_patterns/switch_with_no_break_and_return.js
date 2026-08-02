// vybe-test: js/control_flow_patterns/switch_with_no_break_and_return
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

function classify(n) {
    switch (true) {
        case n < 0: return "negative";
        case n === 0: return "zero";
        case n < 10: return "small";
        default: return "large";
    }
}
__check(__line(classify(-5)), "negative");
__check(__line(classify(0)), "zero");
__check(__line(classify(7)), "small");
__check(__line(classify(100)), "large");
