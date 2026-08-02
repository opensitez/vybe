// vybe-test: js/tagged_template_literal_raw_strings/test_js_tagged_template_higher_order_tag_factory
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

function createPrefixTag(prefix) {
    return (strings, ...values) => {
        return prefix + strings[0] + values[0];
    };
}
const customTag = createPrefixTag("[LOG] ");
__check(__line(customTag`Value = ${100}`), "[LOG] Value = 100");
