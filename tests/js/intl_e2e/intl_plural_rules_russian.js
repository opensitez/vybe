// vybe-test: js/intl_e2e/intl_plural_rules_russian
// origin: languages/js/tests/js/test_intl_e2e.rs

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

const pr = new Intl.PluralRules("ru");
        __check(__line(pr.select(2)), "few");
