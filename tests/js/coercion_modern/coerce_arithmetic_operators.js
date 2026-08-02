// vybe-test: js/coercion_modern/coerce_arithmetic_operators
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

__check(__line("5" - 2), "3");
__check(__line("5" * 2), "10");
__check(__line("5" / 2), "2.5");
__check(__line("5" % 2), "1");
__check(__line("5" + 2), "52");
