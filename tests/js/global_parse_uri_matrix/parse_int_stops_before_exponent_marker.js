// vybe-test: js/global_parse_uri_matrix/parse_int_stops_before_exponent_marker
// origin: languages/js/tests/js/test_global_parse_uri_matrix.rs

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

__check(__line(parseInt("15e2", 10)), "15");
