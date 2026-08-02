// vybe-test: js/array_from_patterns/array_from_set
// origin: languages/js/tests/js/test_array_from_patterns.rs

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

const s = new Set([1, 2, 3, 2, 1]);
const arr = Array.from(s);
__check(__line(arr.join(",")), "1,2,3");
