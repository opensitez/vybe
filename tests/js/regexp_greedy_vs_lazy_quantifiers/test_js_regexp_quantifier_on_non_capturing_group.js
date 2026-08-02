// vybe-test: js/regexp_greedy_vs_lazy_quantifiers/test_js_regexp_quantifier_on_non_capturing_group
// origin: languages/js/tests/js/test_js_regexp_greedy_vs_lazy_quantifiers.rs

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

const str = "ha-ha-ha";
__check(__line(str.match(/(?:ha-)+/)[0]), "ha-ha-");
