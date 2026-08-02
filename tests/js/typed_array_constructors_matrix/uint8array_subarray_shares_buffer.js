// vybe-test: js/typed_array_constructors_matrix/uint8array_subarray_shares_buffer
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

const a=new Uint8Array([1,2,3,4]); const s=a.subarray(1,3); s[0]=9; __check(__line(a[1]), "9");
