// vybe-test: js/typed_arrays/typedarray_from_static_method
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

const a = Int16Array.from([1, 2, 3, 4]);
__check(__line(a.length), "4");
__check(__line(a[0]), "1");
__check(__line(a[3]), "4");
