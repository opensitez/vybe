// vybe-test: js/intl_datetimeformat_timezone_calendar_options/test_js_intl_datetimeformat_invalid_date_input_throws_rangeerror
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

const invalidDate = new Date(NaN);
const formatter = new Intl.DateTimeFormat("en-US");
try {
    formatter.format(invalidDate);
} catch (e) {
    __check(__line("Invalid Date Format RangeError"), "Invalid Date Format RangeError");
}
