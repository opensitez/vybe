// vybe-test: js/date_utc_matrix/date_utc_constructor_accepts_only_year_month
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

const d = new Date(Date.UTC(2024, 0));
__check(__line(d.getUTCDate()), "1");
__check(__line(d.getUTCHours()), "0");
