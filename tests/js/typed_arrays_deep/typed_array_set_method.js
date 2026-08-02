// vybe-test: js/typed_arrays_deep/typed_array_set_method
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

const dest = new Int32Array(8);
const src = [1, 2, 3];
dest.set(src, 2);
__check(__line(dest[0]), "0");
__check(__line(dest[2]), "1");
__check(__line(dest[3]), "2");
