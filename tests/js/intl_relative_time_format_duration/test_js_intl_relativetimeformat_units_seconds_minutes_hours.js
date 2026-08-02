// vybe-test: js/intl_relative_time_format_duration/test_js_intl_relativetimeformat_units_seconds_minutes_hours
// origin: languages/js/tests/js/test_js_intl_relative_time_format_duration.rs

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

const rtf = new Intl.RelativeTimeFormat("en", { numeric: "always" });
__check(__line(`${rtf.format(-30, "second")}:${rtf.format(5, "minute")}:${rtf.format(-2, "hour")}`), "30 seconds ago:in 5 minutes:2 hours ago");
