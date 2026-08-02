// vybe-test: js/ecma_arrays/array_entries_pattern
// origin: languages/js/tests/js/test_ecma_arrays.rs

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
let result = [];
for (let i = 0; i < arr.length; i++) {
    result.push(i + ":" + arr[i]);
}
console.log(result.join(","));
