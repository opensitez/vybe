// vybe-test: js/regex_comprehensive/regex_quantifiers_greedy_lazy
// origin: languages/js/tests/js/test_regex_comprehensive.rs

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

const html = "<b>bold</b> and <i>italic</i>";
const greedy = html.match(/<.+>/)[0];
const lazy = html.match(/<.+?>/)[0];
__check(__line(greedy), "<b>bold</b> and <i>italic</i>");
__check(__line(lazy), "<b>");
