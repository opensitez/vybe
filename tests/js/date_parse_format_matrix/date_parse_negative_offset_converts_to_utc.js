// vybe-test: js/date_parse_format_matrix/date_parse_negative_offset_converts_to_utc
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

const d = new Date(Date.parse("2024-01-02T03:04:05-05:30"));
__check(__line(d.toISOString()), "2024-01-02T08:34:05.000Z");
