// vybe-test: js/date_parse_format_matrix/date_tojson_exact_known_value
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

__check(__line(new Date(Date.UTC(2024, 0, 2, 3, 4, 5)).toJSON()), "2024-01-02T03:04:05.000Z");
