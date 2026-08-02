// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_uint16array_big_endian_data
// origin: languages/js/tests/js/test_js_structured_clone_typed_arrays_array_buffers.rs

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

const u16 = new Uint16Array([0x1234, 0x5678]);
const clone = structuredClone(u16);
__check(__line(clone[0].toString(16) + "|" + clone[1].toString(16)), "1234|5678");
