// vybe-test: js/typed_arrays/int8array_from_array_literal
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

const a = new Int8Array([10, 20, 30]);
__check(__line(a[0]), "10");
__check(__line(a[1]), "20");
__check(__line(a[2]), "30");
