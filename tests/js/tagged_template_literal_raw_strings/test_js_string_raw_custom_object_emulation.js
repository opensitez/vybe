// vybe-test: js/tagged_template_literal_raw_strings/test_js_string_raw_custom_object_emulation
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

const fakeStrings = { raw: ["A", "B"] };
__check(__line(String.raw(fakeStrings, 100)), "A100B");
