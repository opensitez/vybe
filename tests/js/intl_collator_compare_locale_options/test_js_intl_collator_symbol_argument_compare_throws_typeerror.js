// vybe-test: js/intl_collator_compare_locale_options/test_js_intl_collator_symbol_argument_compare_throws_typeerror
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

const collator = new Intl.Collator("en");
try {
    collator.compare(Symbol("a"), "a");
} catch (e) {
    __check(__line("Collator Symbol Argument TypeError"), "Collator Symbol Argument TypeError");
}
