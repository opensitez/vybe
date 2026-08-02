// vybe-test: js/typed_arrays_deep/typed_array_creation_methods
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

const a = new Int32Array([1, 2, 3, 4, 5]);
const b = Int32Array.of(10, 20, 30);
const c = Int32Array.from([1.5, 2.7, 3.9]);
__check(__line(a.length), "5");
__check(__line(b[1]), "20");
__check(__line(c[0]), "1");
