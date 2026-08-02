// vybe-test: js/tagged_templates/tag_no_expressions_gets_one_string
// origin: languages/js/tests/js/test_tagged_templates.rs

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
    __check(__line(strings.length), "1");
    __check(__line(values.length), "0");
    __check(__line(strings[0]), "just a string");
}
tag`just a string`;
