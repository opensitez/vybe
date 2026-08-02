// vybe-test: js/error_handling_advanced/stack_overflow_detection
// origin: languages/js/tests/js/test_error_handling_advanced.rs

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

function recurse(n) {
    try { return recurse(n + 1); }
    catch(e) { return n; }
}
const depth = recurse(0);
__check(__line(depth > 100), "true");
__check(__line(typeof depth), "number");
