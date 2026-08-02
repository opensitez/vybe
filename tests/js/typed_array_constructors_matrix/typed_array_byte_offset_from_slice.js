// vybe-test: js/typed_array_constructors_matrix/typed_array_byte_offset_from_slice
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

const a=new Uint8Array(new ArrayBuffer(4),1,2); __check(__line(a.byteOffset), "1");__check(__line(a.byteLength), "2");
