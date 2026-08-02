// vybe-test: js/closure_scope_deep/mutual_recursion_even_odd
// origin: languages/js/tests/js/test_closure_scope_deep.rs

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

function isEven(n) {
    if (n === 0) return true;
    return isOdd(n - 1);
}
function isOdd(n) {
    if (n === 0) return false;
    return isEven(n - 1);
}
__check(__line(isEven(4)), "true");
__check(__line(isOdd(7)), "true");
__check(__line(isEven(1)), "false");
