// vybe-test: js/regex_patterns_deep/quantifier_greedy_vs_lazy
// origin: languages/js/tests/js/test_regex_patterns_deep.rs

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

const text = "<a><b><c>";
const greedy = text.match(/<.*>/);
const lazy = text.match(/<.*?>/);
__check(__line(greedy[0]), "<a><b><c>");
__check(__line(lazy[0]), "<a>");
