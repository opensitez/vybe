// vybe-test: js/function_edge_cases/iife_creates_isolated_scope
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

const result = (function() {
    const secret = 42;
    return secret * 2;
})();
__check(__line(result), "84");
__check(__line(typeof secret), "undefined");
