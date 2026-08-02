// vybe-test: js/regexp_greedy_vs_lazy_quantifiers/test_js_regexp_invalid_range_min_greater_than_max_throws_syntaxerror
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

try {
    eval("const re = /a{5,2}/;"); // min 5 > max 2 is a SyntaxError!
} catch (e) {
    __check(__line("Invalid Quantifier Range SyntaxError"), "Invalid Quantifier Range SyntaxError");
}
