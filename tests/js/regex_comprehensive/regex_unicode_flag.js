// vybe-test: js/regex_comprehensive/regex_unicode_flag
// origin: languages/js/tests/js/test_regex_comprehensive.rs

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

const emoji = "Hello 😀 World 🌍";
const emojiCount = (emoji.match(/\p{Emoji}/gu) || []).length;
__check(__line(emojiCount >= 2), "true");
const wordCount = "hello world".match(/\p{L}+/gu).length;
__check(__line(wordCount), "2");
