// vybe-test: js/typed_arrays/shared_arraybuffer_between_two_typed_arrays
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

const buf = new ArrayBuffer(8);
const i32 = new Int32Array(buf);
const u8  = new Uint8Array(buf);
i32[0] = 1;
__check(__line(u8[0]), "1");
i32[0] = 256;
__check(__line(u8[0]), "0");
__check(__line(u8[1]), "1");
