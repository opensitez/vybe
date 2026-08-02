// vybe-test: js/regex_string_methods/unicode_flag_handles_surrogates
// origin: languages/js/tests/js/test_regex_string_methods.rs

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

const re = /./u;
const emoji = "😀";
const match = re.exec(emoji);
// With /u, . matches the whole code point (surrogate pair)
__check(__line(match[0].length), "2");
