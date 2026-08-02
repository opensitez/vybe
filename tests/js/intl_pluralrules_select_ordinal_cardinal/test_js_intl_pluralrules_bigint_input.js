// vybe-test: js/intl_pluralrules_select_ordinal_cardinal/test_js_intl_pluralrules_bigint_input
// origin: languages/js/tests/js/test_js_intl_pluralrules_select_ordinal_cardinal.rs

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

const rules = new Intl.PluralRules("en-US");
__check(__line(rules.select(100n)), "other");
