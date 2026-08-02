// vybe-test: js/tagged_templates/tag_raw_different_from_cooked_for_escape
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

function tag(strings) {
    const cooked = strings[0];
    const raw = strings.raw[0];
    __check(__line(cooked !== raw), "true");
}
tag`\n`;
