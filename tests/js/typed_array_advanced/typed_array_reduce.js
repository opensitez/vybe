// vybe-test: js/typed_array_advanced/typed_array_reduce
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

const ta = new Float64Array([1.5, 2.5, 3.5]);
const sum = ta.reduce((acc, x) => acc + x, 0);
__check(__line(sum), "7.5");
