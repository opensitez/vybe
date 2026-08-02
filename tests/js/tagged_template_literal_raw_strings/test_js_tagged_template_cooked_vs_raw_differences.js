// vybe-test: js/tagged_template_literal_raw_strings/test_js_tagged_template_cooked_vs_raw_differences
// origin: languages/js/tests/js/test_js_tagged_template_literal_raw_strings.rs

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

function tag(strings) {
    return strings[0].length + "|" + strings.raw[0].length;
}
__check(__line(tag`\n`), "1|2");
