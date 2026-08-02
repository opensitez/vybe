// vybe-test: js/array_algorithms/longest_common_prefix
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

function longestCommonPrefix(strs) {
    if (!strs.length) return "";
    return strs.reduce((prefix, str) => {
        while (!str.startsWith(prefix)) prefix = prefix.slice(0, -1);
        return prefix;
    });
}
console.log(longestCommonPrefix(["flower", "flow", "flight"]));
console.log(longestCommonPrefix(["dog", "racecar", "car"]));
console.log(longestCommonPrefix(["abc", "abcd", "ab"]));
