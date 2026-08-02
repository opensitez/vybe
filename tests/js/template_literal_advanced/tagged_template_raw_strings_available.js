// vybe-test: js/template_literal_advanced/tagged_template_raw_strings_available
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

function raw(strings) {
    return strings.raw[0];
}
const result = raw`\n\t`;
__check(__line(result.length), "4"); // raw: 4 chars
__check(__line(result[0]), "\\");     // backslash
