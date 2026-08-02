// vybe-test: js/typed_arrays/typedarray_join_method
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

const a = new Int32Array([1, 2, 3, 4]);
__check(__line(a.join(",")), "1,2,3,4");
__check(__line(a.join("-")), "1-2-3-4");
