// vybe-test: js/date_utc_matrix/date_local_setmonth_overflow_rolls_year
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

const d = new Date(2024, 11, 1);
d.setMonth(12);
__check(__line(d.getFullYear()), "2025");
__check(__line(d.getMonth()), "0");
