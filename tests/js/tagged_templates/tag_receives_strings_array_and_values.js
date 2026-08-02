// vybe-test: js/tagged_templates/tag_receives_strings_array_and_values
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
    __check(__line(strings[0]), "before");
    __check(__line(strings[1]), "after");
    __check(__line(values[0]), "42");
}
const x = 42;
tag`before${x}after`;
