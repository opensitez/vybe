// vybe-test: js/typed_arrays/uint32array_length_property
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

const a = new Uint32Array(5);
__check(__line(a.length), "5");
a[4] = 99;
__check(__line(a[4]), "99");
