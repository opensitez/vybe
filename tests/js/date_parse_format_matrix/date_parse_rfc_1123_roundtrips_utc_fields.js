// vybe-test: js/date_parse_format_matrix/date_parse_rfc_1123_roundtrips_utc_fields
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

const d = new Date(Date.parse("Tue, 02 Jan 2024 03:04:05 GMT"));
__check(__line(d.getUTCFullYear()), "2024");
__check(__line(d.getUTCMonth()), "0");
__check(__line(d.getUTCDate()), "2");
__check(__line(d.getUTCHours()), "3");
