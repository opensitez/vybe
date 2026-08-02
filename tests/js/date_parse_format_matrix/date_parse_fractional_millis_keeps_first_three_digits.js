// vybe-test: js/date_parse_format_matrix/date_parse_fractional_millis_keeps_first_three_digits
// origin: languages/js/tests/js/test_date_parse_format_matrix.rs

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

const d = new Date(Date.parse("2024-01-02T03:04:05.6789Z"));
__check(__line(d.getUTCMilliseconds()), "678");
