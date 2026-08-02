// vybe-test: js/tagged_template_literal_raw_strings/test_js_tagged_template_basic_string_array_and_expressions
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

function tag(strings, ...values) {
    return strings[0] + values[0] + strings[1] + values[1] + strings[2];
}
const a = 10, b = 20;
__check(__line(tag`X: ${a}, Y: ${b}!`), "X: 10, Y: 20!");
