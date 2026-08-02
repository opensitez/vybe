// vybe-test: js/error_handling_patterns/error_boundary_pattern
// origin: languages/js/tests/js/test_error_handling_patterns.rs

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

function safe(fn) {
    return function(...args) {
        try {
            return { value: fn(...args), error: null };
        } catch (e) {
            return { value: null, error: e };
        }
    };
}
const safeDivide = safe((a, b) => {
    if (b === 0) throw new Error("division by zero");
    return a / b;
});
const r1 = safeDivide(10, 2);
const r2 = safeDivide(10, 0);
__check(__line(r1.value), "5");
__check(__line(r1.error), "null");
__check(__line(r2.value), "null");
__check(__line(r2.error.message), "division by zero");
