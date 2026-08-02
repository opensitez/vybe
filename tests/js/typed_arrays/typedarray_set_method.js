// vybe-test: js/typed_arrays/typedarray_set_method
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

const a = new Int32Array(5);
a.set([1, 2, 3], 1);
__check(__line(a[0]), "0");
__check(__line(a[1]), "1");
__check(__line(a[2]), "2");
__check(__line(a[3]), "3");
