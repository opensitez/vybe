// vybe-test: js/features/test_recursion
// origin: languages/js/tests/js/js_features_test.rs

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

function fact(n) { if (n <= 1) { return 1; } return n * fact(n - 1); } __check(__line(fact(6)), "720")
