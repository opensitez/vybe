// vybe-test: js/coercion_modern/coerce_string_to_number
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

__check(__line(Number("42")), "42");
__check(__line(Number("3.14")), "3.14");
__check(__line(Number("")), "0");
__check(__line(Number(" ")), "0");
__check(__line(Number("hello")), "NaN");
__check(__line(Number(true)), "1");
__check(__line(Number(false)), "0");
__check(__line(Number(null)), "0");
