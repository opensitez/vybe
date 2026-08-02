// vybe-test: js/coercion_modern/number_coercion_of_arrays
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

__check(__line(Number([])), "0");
__check(__line(Number([5])), "5");
__check(__line(Number([1, 2])), "NaN");
