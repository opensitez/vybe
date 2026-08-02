// vybe-test: js/typed_array_advanced/float32array_precision_loss
// origin: languages/js/tests/js/test_typed_array_advanced.rs

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

const f32 = new Float32Array(1);
f32[0] = Math.PI;
// float32 loses precision compared to float64
__check(__line(f32[0] !== Math.PI), "true");
__check(__line(Math.abs(f32[0] - Math.PI) < 0.001), "true");
