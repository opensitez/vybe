// vybe-test: js/iterator_patterns_deep/entries_iterator_pattern
// origin: languages/js/tests/js/test_iterator_patterns_deep.rs

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

const obj = { a: 1, b: 2, c: 3 };
const entries = Object.entries(obj);
for (const [key, val] of entries) {
    if (key === "b") { console.log(val); break; }
}
// Array entries
const arr = ["x", "y", "z"];
for (const [i, v] of arr.entries()) {
    if (i === 1) { console.log(v); break; }
}
