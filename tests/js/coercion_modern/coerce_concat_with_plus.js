// vybe-test: js/coercion_modern/coerce_concat_with_plus
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

__check(__line("" + 42), "42");
__check(__line("" + true), "true");
__check(__line("" + null), "null");
__check(__line("" + undefined), "undefined");
__check(__line("" + [1,2,3]), "1,2,3");
