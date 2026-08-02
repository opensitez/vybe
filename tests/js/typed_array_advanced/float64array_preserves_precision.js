// vybe-test: js/typed_array_advanced/float64array_preserves_precision
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

const f64 = new Float64Array(1);
f64[0] = Math.PI;
__check(__line(f64[0] === Math.PI), "true");
