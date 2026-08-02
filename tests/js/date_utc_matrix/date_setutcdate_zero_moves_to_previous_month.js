// vybe-test: js/date_utc_matrix/date_setutcdate_zero_moves_to_previous_month
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

const d = new Date(Date.UTC(2024, 2, 1));
d.setUTCDate(0);
__check(__line(d.getUTCMonth()), "1");
__check(__line(d.getUTCDate()), "29");
