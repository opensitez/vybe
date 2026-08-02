// vybe-test: js/global_parse_uri_matrix/isfinite_coerces_null_to_zero
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

__check(__line(isFinite(null)), "true");
