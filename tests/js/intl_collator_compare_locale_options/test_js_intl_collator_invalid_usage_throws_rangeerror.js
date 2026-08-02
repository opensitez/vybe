// vybe-test: js/intl_collator_compare_locale_options/test_js_intl_collator_invalid_usage_throws_rangeerror
// origin: languages/js/tests/js/test_js_intl_collator_compare_locale_options.rs

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
    new Intl.Collator("en", { usage: "invalid_usage" });
} catch (e) {
    __check(__line("Invalid Usage RangeError"), "Invalid Usage RangeError");
}
