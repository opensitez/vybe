// vybe-test: js/intl_datetimeformat_timezone_calendar_options/test_js_intl_datetimeformat_timestamp_number_input
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

const timestamp = 1784678400000; // Epoch timestamp
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", year: "numeric" });
__check(__line(formatter.format(timestamp)), "2026");
