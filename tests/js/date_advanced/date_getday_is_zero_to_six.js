// vybe-test: js/date_advanced/date_getday_is_zero_to_six
// origin: languages/js/tests/js/test_date_advanced.rs

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

// 2024-01-07 is a Sunday (0)
const d = new Date("2024-01-07T12:00:00.000Z");
__check(__line(d.getUTCDay()), "0");
