// vybe-test: js/array_es2023/array_from_on_set
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

const s = new Set([1, 2, 2, 3, 3, 3]);
const arr = Array.from(s);
arr.sort((a, b) => a - b);
__check(__line(arr.join(",")), "1,2,3");
