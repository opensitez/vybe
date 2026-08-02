// vybe-test: js/typed_arrays/bigint64array_creation_with_bigint_values
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

const a = new BigInt64Array([1n, 2n, 9007199254740993n]);
__check(__line(a[0]), "1n");
__check(__line(a[1]), "2n");
__check(__line(a[2]), "9007199254740993n");
__check(__line(a.length), "3");
