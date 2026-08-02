// vybe-test: js/number_precision_edge/parseint_stops_at_invalid_character
// origin: languages/js/tests/js/test_number_precision_edge.rs

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

__check(__line(parseInt("10.5")), "10");
__check(__line(parseInt("0xFF")), "255");
__check(__line(parseInt("")), "NaN");
__check(__line(isNaN(parseInt(""))), "true");
