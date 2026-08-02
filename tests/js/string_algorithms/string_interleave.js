// vybe-test: js/string_algorithms/string_interleave
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

function interleave(a, b) {
    const result = [];
    const len = Math.max(a.length, b.length);
    for (let i = 0; i < len; i++) {
        if (i < a.length) result.push(a[i]);
        if (i < b.length) result.push(b[i]);
    }
    return result.join("");
}
console.log(interleave("abc", "12345"));
console.log(interleave("xyz", ""));
