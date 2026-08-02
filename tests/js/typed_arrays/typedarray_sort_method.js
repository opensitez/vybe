// vybe-test: js/typed_arrays/typedarray_sort_method
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

const a = new Int32Array([3, 1, 4, 1, 5, 9]);
a.sort();
__check(__line(a.join(",")), "1,1,3,4,5,9");
