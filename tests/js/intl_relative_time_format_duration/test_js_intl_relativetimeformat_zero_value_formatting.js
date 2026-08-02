// vybe-test: js/intl_relative_time_format_duration/test_js_intl_relativetimeformat_zero_value_formatting
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

const rtfAlways = new Intl.RelativeTimeFormat("en", { numeric: "always" });
__check(__line(rtfAlways.format(0, "second") + "|" + rtfAlways.format(-0, "second")), "in 0 seconds|0 seconds ago");
