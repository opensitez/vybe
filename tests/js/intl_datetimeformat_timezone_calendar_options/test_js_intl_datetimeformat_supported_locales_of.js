// vybe-test: js/intl_datetimeformat_timezone_calendar_options/test_js_intl_datetimeformat_supported_locales_of
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

const locales = Intl.DateTimeFormat.supportedLocalesOf(["en-US", "ja-JP"]);
__check(__line(locales.length), "2");
