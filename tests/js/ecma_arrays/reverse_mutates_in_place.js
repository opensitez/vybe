// vybe-test: js/ecma_arrays/reverse_mutates_in_place
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

const arr = [1, 2, 3];
const same = arr.reverse();
__check(__line(arr === same), "true");
__check(__line(arr.join(",")), "3,2,1");
