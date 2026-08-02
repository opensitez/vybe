// vybe-test: js/typed_arrays/typedarray_slice_creates_copy
// origin: languages/js/tests/js/test_typed_arrays.rs

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

const a = new Int32Array([10, 20, 30, 40, 50]);
const copy = a.slice(1, 4);
__check(__line(copy.length), "3");
__check(__line(copy[0]), "20");
__check(__line(copy[2]), "40");
a[1] = 999;
__check(__line(copy[0]), "20");
