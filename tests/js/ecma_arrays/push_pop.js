// vybe-test: js/ecma_arrays/push_pop
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

const arr = [1, 2];
arr.push(3);
__check(__line(arr.length), "3");
const last = arr.pop();
__check(__line(last), "3");
__check(__line(arr.length), "2");
