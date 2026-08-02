// vybe-test: js/string_unicode/from_codepoint_emoji
// origin: languages/js/tests/js/test_string_unicode.rs

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

const emoji = String.fromCodePoint(0x1F600);
__check(__line(emoji.length), "2");
__check(__line(emoji.charCodeAt(0).toString(16)), "d83d"); // high surrogate of U+1F600
