// vybe-test: js/coercion_modern/coerce_double_bang
// origin: languages/js/tests/js/test_coercion_modern.rs

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

__check(__line(!!0), "false");
__check(__line(!!""), "false");
__check(__line(!!null), "false");
__check(__line(!!1), "true");
__check(__line(!!"hi"), "true");
__check(__line(!!{}), "true");
