// vybe-test: js/destructuring_array_deep/destructure_map_entries_pattern
// origin: languages/js/tests/js/test_destructuring_array_deep.rs

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

const map = new Map([["x", 1], ["y", 2]]);
for (const [key, val] of map) {
    console.log(key + "=" + val);
}
