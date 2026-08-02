// vybe-test: js/template_literal_advanced/tagged_template_strings_array_length
// origin: languages/js/tests/js/test_template_literal_advanced.rs

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
    // strings.length is always values.length + 1
    __check(__line(strings.length), "3");
    __check(__line(values.length), "2");
}
tag`a${1}b${2}c`;
