// vybe-test: js/array_es2023/array_es2023_with_negative_index_support
// origin: languages/js/tests/js/test_array_es2023.rs

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

const arr = [10, 20, 30];
const updated = arr.with(-1, 99);
__check(__line(arr.join(",") + "|" + updated.join(",")), "10,20,30|10,20,99");
