// vybe-test: js/string_algorithms/reverse_words
// origin: languages/js/tests/js/test_string_algorithms.rs

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

function reverseWords(s) {
    return s.trim().split(/\s+/).reverse().join(" ");
}
__check(__line(reverseWords("Hello World")), "World Hello");
__check(__line(reverseWords("  the sky is blue  ")), "blue is sky the");
