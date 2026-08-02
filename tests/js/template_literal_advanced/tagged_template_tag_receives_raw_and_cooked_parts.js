// vybe-test: js/template_literal_advanced/tagged_template_tag_receives_raw_and_cooked_parts
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

function inspect(strings, value) {
    __check(__line(strings.length), "2");
    __check(__line(strings[0].length), "2");
    __check(__line(strings.raw[0].length), "3");
    __check(__line(value), "B");
}
inspect`a\n${"B"}\t`;
