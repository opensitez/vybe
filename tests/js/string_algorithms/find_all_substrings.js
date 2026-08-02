// vybe-test: js/string_algorithms/find_all_substrings
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

function findAll(text, pattern) {
    const indices = [];
    let idx = text.indexOf(pattern);
    while (idx !== -1) {
        indices.push(idx);
        idx = text.indexOf(pattern, idx + 1);
    }
    return indices;
}
console.log(findAll("abababab", "ab").join(","));
console.log(findAll("hello", "xyz").join(","));
