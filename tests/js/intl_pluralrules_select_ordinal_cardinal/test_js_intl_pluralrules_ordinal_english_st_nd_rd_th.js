// vybe-test: js/intl_pluralrules_select_ordinal_cardinal/test_js_intl_pluralrules_ordinal_english_st_nd_rd_th
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

const rules = new Intl.PluralRules("en-US", { type: "ordinal" });
__check(__line(`${rules.select(1)}:${rules.select(2)}:${rules.select(3)}:${rules.select(4)}:${rules.select(11)}`), "one:two:few:other:other");
