// vybe-test: js/ecma_arrays/shift_unshift
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

const arr = [2, 3];
arr.unshift(1);
__check(__line(arr.join(",")), "1,2,3");
const first = arr.shift();
__check(__line(first), "1");
__check(__line(arr.join(",")), "2,3");
