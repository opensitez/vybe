// vybe-test: js/function_edge_cases/default_param_can_reference_earlier_param
// origin: languages/js/tests/js/test_function_edge_cases.rs

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

function rect(w, h = w) { return w * h; }
__check(__line(rect(5)), "25");
__check(__line(rect(3, 4)), "12");
