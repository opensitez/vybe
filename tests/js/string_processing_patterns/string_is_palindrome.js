// vybe-test: js/string_processing_patterns/string_is_palindrome
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

function isPalindrome(str) {
    const clean = str.toLowerCase().replace(/[^a-z0-9]/g, "");
    return clean === [...clean].reverse().join("");
}
__check(__line(isPalindrome("racecar")), "true");
__check(__line(isPalindrome("A man a plan a canal Panama")), "true");
__check(__line(isPalindrome("hello")), "false");
