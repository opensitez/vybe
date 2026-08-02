// vybe-test: js/string_algorithms/count_vowels
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

const countVowels = s => (s.match(/[aeiouAEIOU]/g) || []).length;
__check(__line(countVowels("Hello World")), "3");
__check(__line(countVowels("rhythm")), "0");
__check(__line(countVowels("aeiou")), "5");
