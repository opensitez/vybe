// vybe-test: js/string_processing_patterns/count_occurrences
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

function count(str, sub) {
    let n = 0, pos = 0;
    while ((pos = str.indexOf(sub, pos)) !== -1) { n++; pos += sub.length; }
    return n;
}
console.log(count("abcabcabc", "abc"));
console.log(count("hello world hello", "hello"));
console.log(count("aaa", "aa")); // non-overlapping
