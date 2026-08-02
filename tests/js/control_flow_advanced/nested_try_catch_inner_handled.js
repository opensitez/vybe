// vybe-test: js/control_flow_advanced/nested_try_catch_inner_handled
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

let result = "none";
try {
    try { throw new RangeError("inner"); }
    catch (e) {
        if (e instanceof RangeError) result = "range";
        else throw e;
    }
} catch (e) {
    result = "outer";
}
__check(__line(result), "range");
