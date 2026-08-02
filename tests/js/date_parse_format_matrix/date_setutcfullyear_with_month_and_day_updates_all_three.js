// vybe-test: js/date_parse_format_matrix/date_setutcfullyear_with_month_and_day_updates_all_three
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

const d = new Date(Date.UTC(2024, 0, 1));
d.setUTCFullYear(2025, 5, 15);
__check(__line(d.getUTCFullYear()), "2025");
__check(__line(d.getUTCMonth()), "5");
__check(__line(d.getUTCDate()), "15");
