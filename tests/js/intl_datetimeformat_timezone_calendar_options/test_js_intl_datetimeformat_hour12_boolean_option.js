// vybe-test: js/intl_datetimeformat_timezone_calendar_options/test_js_intl_datetimeformat_hour12_boolean_option
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

const date = new Date(Date.UTC(2026, 0, 1, 14, 30, 0));
const fmt12 = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", hour: "numeric", hour12: true });
const fmt24 = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", hour: "numeric", hour12: false });
__check(__line(fmt12.format(date) + "|" + fmt24.format(date)), "2 PM|14");
