// vybe-test: js/typed_array_constructors_matrix/bigint64array_subarray_preserves_element_type
// origin: languages/js/tests/js/test_typed_array_constructors_matrix.rs

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

const a=new BigInt64Array([1n,2n,3n]); __check(__line(typeof a.subarray(1)[0]), "bigint");
