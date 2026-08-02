// vybe-test: js/intl_datetimeformat_timezone_calendar_options/test_js_intl_datetimeformat_format_range_to_parts
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

const d1 = new Date(Date.UTC(2026, 6, 20));
const d2 = new Date(Date.UTC(2026, 6, 22));
const formatter = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", day: "numeric" });
const parts = formatter.formatRangeToParts(d1, d2);
__check(__line(parts.some(p => p.type === "day")), "true");
