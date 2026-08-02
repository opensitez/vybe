// vybe-test: js/numeric_coercion_deep/bitwise_coerces_to_int32
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

__check(__line(3.7 | 0), "3");
__check(__line(-3.7 | 0), "-3");
__check(__line(2**32 + 1 | 0), "1"); // wraps around int32
