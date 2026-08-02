// vybe-test: js/intl_datetimeformat_timezone_calendar_options/test_js_intl_datetimeformat_numeric_vs_2digit_day
// origin: languages/js/tests/js/test_js_intl_datetimeformat_timezone_calendar_options.rs

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

const date = new Date(Date.UTC(2026, 6, 5));
const fmtNum = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", day: "numeric" });
const fmt2Digit = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", day: "2-digit" });
__check(__line(fmtNum.format(date) + "|" + fmt2Digit.format(date)), "5|05");
