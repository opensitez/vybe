// vybe-test: js/closure_scope_deep_patterns/partial_application_closure
// origin: languages/js/tests/js/test_closure_scope_deep_patterns.rs

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

function multiply(a) {
    return function(b) {
        return function(c) {
            return a * b * c;
        };
    };
}
const double = multiply(2);
const times6 = double(3);
__check(__line(times6(4)), "24");
__check(__line(double(5)(6)), "60");
