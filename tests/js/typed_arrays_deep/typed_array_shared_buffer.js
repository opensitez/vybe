// vybe-test: js/typed_arrays_deep/typed_array_shared_buffer
// origin: languages/js/tests/js/test_typed_arrays_deep.rs

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

const buffer = new ArrayBuffer(16);
const int32 = new Int32Array(buffer);
int32[0] = 1;
int32[1] = 256;
__check(__line(int32[0]), "1");
__check(__line(int32[1]), "256");
__check(__line(buffer.byteLength), "16");
