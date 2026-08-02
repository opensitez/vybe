// vybe-test: js/numeric_coercion_deep/parse_float_basic
// origin: languages/js/tests/js/test_numeric_coercion_deep.rs

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

__check(__line(parseFloat("3.14")), "3.14");
__check(__line(parseFloat(".5")), "0.5");
__check(__line(parseFloat("1e3")), "1000");
__check(__line(parseFloat("Infinity")), "Infinity");
