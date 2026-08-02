// vybe-test: js/function_edge_cases/method_shorthand_in_object_literal
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

const obj = {
    double(x) { return x * 2; },
    triple(x) { return x * 3; }
};
__check(__line(obj.double(5)), "10");
__check(__line(obj.triple(5)), "15");
