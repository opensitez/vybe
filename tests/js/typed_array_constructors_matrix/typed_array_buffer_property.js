// vybe-test: js/typed_array_constructors_matrix/typed_array_buffer_property
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

const buf=new ArrayBuffer(8); const a=new Uint8Array(buf); __check(__line(a.buffer===buf), "true");
