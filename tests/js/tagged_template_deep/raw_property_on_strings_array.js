// vybe-test: js/tagged_template_deep/raw_property_on_strings_array
// origin: languages/js/tests/js/test_tagged_template_deep.rs

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
    return strings.raw[0];
}
const result = tag`\n\t`;
__check(__line(result), "\\n\\t"); // raw: \\n\\t
__check(__line(result.length), "4");
