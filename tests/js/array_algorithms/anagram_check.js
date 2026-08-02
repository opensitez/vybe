// vybe-test: js/array_algorithms/anagram_check
// origin: languages/js/tests/js/test_array_algorithms.rs

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

function isAnagram(a, b) {
    if (a.length !== b.length) return false;
    const sort = s => s.split("").sort().join("");
    return sort(a) === sort(b);
}
__check(__line(isAnagram("listen", "silent")), "true");
__check(__line(isAnagram("hello", "world")), "false");
__check(__line(isAnagram("anagram", "nagaram")), "true");
