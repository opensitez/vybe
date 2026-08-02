// vybe-test: js/intl_collator_compare_locale_options/test_js_intl_collator_resolved_options
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

const collator = new Intl.Collator("en-US", { numeric: true, sensitivity: "base" });
const opts = collator.resolvedOptions();
__check(__line(opts.locale + "|" + opts.numeric + "|" + opts.sensitivity), "en-US|true|base");
