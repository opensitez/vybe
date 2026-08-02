// vybe-test: js/intl_datetimeformat_timezone_calendar_options/test_js_intl_datetimeformat_invalid_timezone_throws_rangeerror
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

try {
    new Intl.DateTimeFormat("en-US", { timeZone: "Invalid/Timezone_Name" });
} catch (e) {
    __check(__line("Invalid TimeZone RangeError"), "Invalid TimeZone RangeError");
}
