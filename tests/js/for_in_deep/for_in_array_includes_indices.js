// vybe-test: js/for_in_deep/for_in_array_includes_indices
// origin: languages/js/tests/js/test_for_in_deep.rs

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

const arr = ["a", "b", "c"];
const keys = [];
for (const k in arr) keys.push(k);
console.log(keys.join(","));
