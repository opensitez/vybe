// vybe-test: js/number_advanced/math_advanced_functions
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

__check(__line(Math.sign(-5)), "-1");
__check(__line(Math.sign(0)), "0");
__check(__line(Math.sign(7)), "1");
__check(__line(Math.hypot(3, 4)), "5");
__check(__line(Math.log2(8)), "3");
__check(__line(Math.log10(1000)), "3");
