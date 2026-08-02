// vybe-test: js/typed_array_advanced/multiple_views_of_same_buffer
// origin: languages/js/tests/js/test_typed_array_advanced.rs

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

const buffer = new ArrayBuffer(8);
const i32 = new Int32Array(buffer);
const u8 = new Uint8Array(buffer);

i32[0] = 1; // sets first 4 bytes
// u8 sees the same memory
__check(__line(u8[0] !== 0 || u8[1] !== 0 || u8[2] !== 0 || u8[3] !== 0), "true");
