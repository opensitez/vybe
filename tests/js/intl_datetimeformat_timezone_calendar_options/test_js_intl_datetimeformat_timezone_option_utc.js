// vybe-test: js/intl_datetimeformat_timezone_calendar_options/test_js_intl_datetimeformat_timezone_option_utc
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

const date = new Date(Date.UTC(2026, 0, 1, 0, 0, 0));
const fmtUTC = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", hour: "numeric", minute: "numeric" });
__check(__line(fmtUTC.format(date)), "12:00 AM");
