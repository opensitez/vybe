// vybe-test: js/typed_arrays/float32array_creation_and_element_access
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

const a = new Float32Array([1.5, 2.5, 3.5]);
__check(__line(a[0]), "1.5");
__check(__line(a[1]), "2.5");
__check(__line(a[2]), "3.5");
__check(__line(a.length), "3");
