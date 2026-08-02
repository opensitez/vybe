// vybe-test: js/dataview_typed_array_deep/uint8clampedarray_clamps_values
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

const arr = new Uint8ClampedArray(3);
arr[0] = 300;  // clamped to 255
arr[1] = -10;  // clamped to 0
arr[2] = 128;  // unchanged
__check(__line(arr[0]), "255");
__check(__line(arr[1]), "0");
__check(__line(arr[2]), "128");
