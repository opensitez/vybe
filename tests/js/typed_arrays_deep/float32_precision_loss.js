// vybe-test: js/typed_arrays_deep/float32_precision_loss
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

const f64 = 1.337;
const arr = new Float32Array(1);
arr[0] = f64;
const f32 = arr[0];
__check(__line(f32 !== f64), "true");
__check(__line(Math.abs(f32 - f64) < 0.0001), "true");
