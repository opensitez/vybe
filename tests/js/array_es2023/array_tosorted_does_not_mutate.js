// vybe-test: js/array_es2023/array_tosorted_does_not_mutate
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

const orig = [3, 1, 4, 1, 5];
const sorted = orig.toSorted((a, b) => a - b);
__check(__line(orig.join(",")), "3,1,4,1,5");
__check(__line(sorted.join(",")), "1,1,3,4,5");
