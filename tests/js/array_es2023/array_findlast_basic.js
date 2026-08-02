// vybe-test: js/array_es2023/array_findlast_basic
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

const arr = [1, 3, 5, 7, 2, 4, 6];
const last = arr.findLast(x => x % 2 === 0);
__check(__line(last), "6");
