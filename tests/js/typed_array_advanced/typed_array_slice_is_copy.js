// vybe-test: js/typed_array_advanced/typed_array_slice_is_copy
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

const ta = new Int32Array([1, 2, 3, 4]);
const sliced = ta.slice(1, 3);
sliced[0] = 99;
__check(__line(ta[1]), "2"); // unchanged
__check(__line(sliced[0]), "99");
