// vybe-test: js/date/date_month_overflow
// origin: languages/js/tests/js/test_date.rs

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

let d = new Date(2024, 11, 31);
d.setMonth(d.getMonth() + 1);
__check(__line(d.getFullYear()), "2025");
__check(__line(d.getMonth()), "0");
