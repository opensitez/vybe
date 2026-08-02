// vybe-test: js/dataview_typed_array_deep/typed_array_shared_buffer
// origin: languages/js/tests/js/test_dataview_typed_array_deep.rs

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
const u8 = new Uint8Array(buf);
i32[0] = 0x01020304;
// u8 sees the bytes of i32[0]
const bytes = [u8[0], u8[1], u8[2], u8[3]];
__check(__line(bytes.some(b => b !== 0)), "true");
