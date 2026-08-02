// vybe-test: js/math_algorithms/number_palindrome
// origin: languages/js/tests/js/test_math_algorithms.rs

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

function isPalindromeNum(n) {
    if (n < 0) return false;
    const s = n.toString();
    return s === s.split("").reverse().join("");
}
__check(__line(isPalindromeNum(121)), "true");
__check(__line(isPalindromeNum(-121)), "false");
__check(__line(isPalindromeNum(1001)), "true");
__check(__line(isPalindromeNum(10)), "false");
