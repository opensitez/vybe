// vybe-test: js/typed_arrays/int8array_overflow_wraps
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

const a = new Int8Array([127, 128, 129, -128, -129]);
__check(__line(a[0]), "127");
__check(__line(a[1]), "-128");
__check(__line(a[2]), "-127");
__check(__line(a[3]), "-128");
__check(__line(a[4]), "127");
