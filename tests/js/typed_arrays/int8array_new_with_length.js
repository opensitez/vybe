// vybe-test: js/typed_arrays/int8array_new_with_length
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

const a = new Int8Array(3);
__check(__line(a.length), "3");
__check(__line(a[0]), "0");
__check(__line(a[1]), "0");
__check(__line(a[2]), "0");
