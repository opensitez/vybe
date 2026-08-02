// vybe-test: js/string_algorithms/most_common_char
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

function mostCommon(s) {
    const freq = {};
    for (const c of s) freq[c] = (freq[c] ?? 0) + 1;
    return Object.entries(freq).sort((a, b) => b[1] - a[1])[0][0];
}
console.log(mostCommon("aabbccddeeee"));
console.log(mostCommon("hello"));
