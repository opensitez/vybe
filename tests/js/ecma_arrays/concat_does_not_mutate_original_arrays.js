// vybe-test: js/ecma_arrays/concat_does_not_mutate_original_arrays
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

const a = [1, 2];
const b = [3, 4];
const c = a.concat(b);
__check(__line(a.join(",")), "1,2");
__check(__line(b.join(",")), "3,4");
__check(__line(c.join(",")), "1,2,3,4");
