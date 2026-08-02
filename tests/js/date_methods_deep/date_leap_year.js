// vybe-test: js/date_methods_deep/date_leap_year
// origin: languages/js/tests/js/test_date_methods_deep.rs

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

const d = new Date("2024-02-29T00:00:00.000Z");
__check(__line(d.getUTCFullYear()), "2024");
__check(__line(d.getUTCMonth()), "1"); // February = 1
__check(__line(d.getUTCDate()), "29");
__check(__line(isNaN(d.getTime())), "false"); // valid date
