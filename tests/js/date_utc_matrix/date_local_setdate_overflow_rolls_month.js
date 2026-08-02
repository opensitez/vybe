// vybe-test: js/date_utc_matrix/date_local_setdate_overflow_rolls_month
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

const d = new Date(2024, 0, 31);
d.setDate(32);
__check(__line(d.getMonth()), "1");
__check(__line(d.getDate()), "1");
