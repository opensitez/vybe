// vybe-test: js/intl_relative_time_format_duration/test_js_intl_relativetimeformat_resolved_options
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

const rtf = new Intl.RelativeTimeFormat("en-US", { numeric: "auto", style: "short" });
const opts = rtf.resolvedOptions();
__check(__line(opts.locale + "|" + opts.numeric + "|" + opts.style), "en-US|auto|short");
