// vybe-test: js/features/test_higher_order_function
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

function apply(f, x) { return f(x); }
        let double = (x) => x * 2;
        __check(__line(apply(double, 21)), "42");
