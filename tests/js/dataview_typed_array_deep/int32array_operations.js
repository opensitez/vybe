// vybe-test: js/dataview_typed_array_deep/int32array_operations
// origin: languages/js/tests/js/test_dataview_typed_array_deep.rs

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

const arr = new Int32Array([1, 2, 3, 4, 5]);
const sum = arr.reduce((acc, x) => acc + x, 0);
__check(__line(sum), "15");
__check(__line(arr.length), "5");
