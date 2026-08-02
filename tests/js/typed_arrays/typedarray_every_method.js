// vybe-test: js/typed_arrays/typedarray_every_method
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

const a = new Int32Array([2, 4, 6, 8]);
__check(__line(a.every(x => x % 2 === 0)), "true");
const b = new Int32Array([2, 3, 6]);
__check(__line(b.every(x => x % 2 === 0)), "false");
