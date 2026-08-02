// vybe-test: js/number_advanced/float_comparison_epsilon
// origin: languages/js/tests/js/test_number_advanced.rs

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

function aboutEqual(a, b, eps = Number.EPSILON) {
    return Math.abs(a - b) <= eps * Math.max(Math.abs(a), Math.abs(b));
}
__check(__line(0.1 + 0.2 === 0.3), "false");
__check(__line(aboutEqual(0.1 + 0.2, 0.3, 1e-10)), "true");
__check(__line(aboutEqual(1.0, 1.0 + 1e-16, 1e-10)), "true");
