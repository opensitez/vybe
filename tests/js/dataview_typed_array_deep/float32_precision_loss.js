// vybe-test: js/dataview_typed_array_deep/float32_precision_loss
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

const f32 = new Float32Array(1);
f32[0] = 1.337;
// Float32 has less precision than Float64
console.log(f32[0] !== 1.337);
// But it's close
console.log(Math.abs(f32[0] - 1.337) < 0.001);
