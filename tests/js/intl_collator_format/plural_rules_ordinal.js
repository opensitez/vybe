// vybe-test: js/intl_collator_format/plural_rules_ordinal
// origin: languages/js/tests/js/test_intl_collator_format.rs

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

const pr = new Intl.PluralRules("en-US", { type: "ordinal" });
__check(__line(pr.select(1)), "one");  // "one" → 1st
__check(__line(pr.select(2)), "two");  // "two" → 2nd
__check(__line(pr.select(3)), "few");  // "few" → 3rd
