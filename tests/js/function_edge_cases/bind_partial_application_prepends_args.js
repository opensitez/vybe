// vybe-test: js/function_edge_cases/bind_partial_application_prepends_args
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

function multiply(a, b) { return a * b; }
const double = multiply.bind(null, 2);
__check(__line(double(5)), "10");
__check(__line(double(10)), "20");
