// vybe-test: js/tagged_template_deep/tagged_template_cooked_array
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
    return strings.join("-");
}
__check(__line(tag`a${1}b${2}c`), "a-b-c");
