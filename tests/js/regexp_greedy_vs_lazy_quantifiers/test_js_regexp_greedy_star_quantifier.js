// vybe-test: js/regexp_greedy_vs_lazy_quantifiers/test_js_regexp_greedy_star_quantifier
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

const html = "<div>first</div><div>second</div>";
__check(__line(html.match(/<div>.*<\/div>/)[0]), "<div>first</div><div>second</div>");
