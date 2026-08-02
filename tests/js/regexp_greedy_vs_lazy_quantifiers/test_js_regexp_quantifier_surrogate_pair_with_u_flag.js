// vybe-test: js/regexp_greedy_vs_lazy_quantifiers/test_js_regexp_quantifier_surrogate_pair_with_u_flag
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

const str = "😀😀😀";
__check(__line(str.match(/😀{2}/u)[0]), "😀😀");
