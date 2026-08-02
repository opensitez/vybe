// vybe-test: js/strict_mode/sloppy_mode_arguments_aliased
// origin: languages/js/tests/js/test_strict_mode.rs

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

function f(a) {
    // sloppy mode — arguments[0] aliases parameter a
    arguments[0] = 99;
    return a;
}
__check(__line(f(1)), "99");
