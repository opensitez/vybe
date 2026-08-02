// vybe-test: js/date_methods_deep/date_get_utc_methods
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

const d = new Date("2024-03-15T10:30:45.500Z");
__check(__line(d.getUTCFullYear()), "2024");
__check(__line(d.getUTCMonth()), "2"); // 2 (March, 0-indexed)
__check(__line(d.getUTCDate()), "15");
__check(__line(d.getUTCHours()), "10");
__check(__line(d.getUTCMinutes()), "30");
__check(__line(d.getUTCSeconds()), "45");
__check(__line(d.getUTCMilliseconds()), "500");
