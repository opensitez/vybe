// vybe-test: js/typed_arrays/uint8clampedarray_clamps_above_255
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

const a = new Uint8ClampedArray([0, 128, 255, 300, 1000]);
__check(__line(a[0]), "0");
__check(__line(a[2]), "255");
__check(__line(a[3]), "255");
__check(__line(a[4]), "255");
