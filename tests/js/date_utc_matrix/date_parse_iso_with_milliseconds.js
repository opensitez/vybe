// vybe-test: js/date_utc_matrix/date_parse_iso_with_milliseconds
// origin: languages/js/tests/js/test_date_utc_matrix.rs

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

const ts = Date.parse("2024-01-02T03:04:05.006Z");
const d = new Date(ts);
__check(__line(d.getUTCSeconds()), "5");
__check(__line(d.getUTCMilliseconds()), "6");
