// vybe-test: js/intl_relative_time_format_duration/test_js_intl_relativetimeformat_symbol_number_argument_throws_typeerror
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

const rtf = new Intl.RelativeTimeFormat("en");
try {
    rtf.format(Symbol("1"), "day");
} catch (e) {
    __check(__line("RelativeTimeFormat Symbol TypeError"), "RelativeTimeFormat Symbol TypeError");
}
