// vybe-test: js/number_edge_cases/infinity_arithmetic
// origin: languages/js/tests/js/test_number_edge_cases.rs

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

__check(__line(Infinity + 1), "Infinity");
__check(__line(Infinity - Infinity), "NaN");
__check(__line(Infinity * -1), "-Infinity");
__check(__line(1 / 0), "Infinity");
__check(__line(-1 / 0), "-Infinity");
