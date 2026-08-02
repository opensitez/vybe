// vybe-test: js/date_utc_matrix/date_utc_negative_month_in_constructor_rolls_previous_year
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

const d = new Date(Date.UTC(2024, -1, 1));
__check(__line(d.getUTCFullYear()), "2023");
__check(__line(d.getUTCMonth()), "11");
