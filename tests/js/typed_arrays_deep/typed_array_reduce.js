// vybe-test: js/typed_arrays_deep/typed_array_reduce
// origin: languages/js/tests/js/test_typed_arrays_deep.rs

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

const arr = new Float64Array([1.5, 2.5, 3.0, 4.0]);
const sum = arr.reduce((a, b) => a + b, 0);
const max = arr.reduce((a, b) => Math.max(a, b), -Infinity);
__check(__line(sum), "11");
__check(__line(max), "4");
