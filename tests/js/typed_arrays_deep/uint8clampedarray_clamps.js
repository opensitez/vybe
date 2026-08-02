// vybe-test: js/typed_arrays_deep/uint8clampedarray_clamps
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

const arr = new Uint8ClampedArray(4);
arr[0] = 300;
arr[1] = -10;
arr[2] = 128;
arr[3] = 0;
__check(__line(arr[0]), "255");
__check(__line(arr[1]), "0");
__check(__line(arr[2]), "128");
__check(__line(arr[3]), "0");
