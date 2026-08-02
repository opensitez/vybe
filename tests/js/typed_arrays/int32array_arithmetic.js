// vybe-test: js/typed_arrays/int32array_arithmetic
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

const a = new Int32Array([10, 20, 30]);
const sum = a[0] + a[1] + a[2];
__check(__line(sum), "60");
a[0] = a[1] * 2;
__check(__line(a[0]), "40");
