// vybe-test: js/intl_relative_time_format_duration/test_js_intl_relativetimeformat_past_and_future_days
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
__check(__line(rtf.format(-1, "day") + "|" + rtf.format(1, "day")), "1 day ago|in 1 day");
