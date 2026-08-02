// vybe-test: js/string_algorithms/compress_string
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

function compress(s) {
    let result = "";
    let i = 0;
    while (i < s.length) {
        let j = i;
        while (j < s.length && s[j] === s[i]) j++;
        result += s[i] + (j - i > 1 ? (j - i) : "");
        i = j;
    }
    return result.length < s.length ? result : s;
}
console.log(compress("aabcccdddd"));
console.log(compress("abc"));
