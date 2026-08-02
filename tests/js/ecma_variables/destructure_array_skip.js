// vybe-test: js/ecma_variables/destructure_array_skip
// origin: languages/js/tests/js/test_ecma_variables.rs

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

const arr = [1, 2, 3, 4];
const [a, , c] = arr;
__check(__line(a), "1");
__check(__line(c), "3");
