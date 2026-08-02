// vybe-test: js/ecma_iterators/for_in_skips_deleted_array_holes
// origin: languages/js/tests/js/test_ecma_iterators.rs

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
delete arr[1];
let keys = [];
for (const key in arr) {
    keys.push(key);
}
console.log(keys.join(","));
