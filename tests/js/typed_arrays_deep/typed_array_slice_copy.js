// vybe-test: js/typed_arrays_deep/typed_array_slice_copy
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

const orig = new Float32Array([1.0, 2.0, 3.0, 4.0]);
const sliced = orig.slice(1, 3);
sliced[0] = 99;
__check(__line(orig[1]), "2");
__check(__line(sliced[0]), "99");
__check(__line(sliced.length), "2");
