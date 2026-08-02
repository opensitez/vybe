// vybe-test: js/string_processing_patterns/word_count
// origin: languages/js/tests/js/test_string_processing_patterns.rs

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

function wordCount(str) {
    return str.trim().split(/\s+/).filter(Boolean).length;
}
__check(__line(wordCount("hello world foo")), "3");
__check(__line(wordCount("  spaced  words  ")), "2");
__check(__line(wordCount("")), "0");
