// vybe-test: js/multiline_template_literals_escapes/test_js_template_literal_octal_escapes_prohibited_in_tagged_cooked
// origin: languages/js/tests/js/test_js_multiline_template_literals_escapes.rs

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
    return strings[0] === undefined; // Cooked is undefined for non-octal legacy escapes in ES2018
}
console.log(tag`\0123`);
