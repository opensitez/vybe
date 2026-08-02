// vybe-test: js/string_unicode_deep/for_of_iterates_code_points
// origin: languages/js/tests/js/test_string_unicode_deep.rs

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

const emoji = "😀"; // actual emoji literal (2 code units, 1 codepoint)
const chars = [...emoji];
__check(__line(chars.length), "1"); // 1 character
