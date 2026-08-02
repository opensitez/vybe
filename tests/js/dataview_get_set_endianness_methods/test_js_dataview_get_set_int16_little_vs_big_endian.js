// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_get_set_int16_little_vs_big_endian
// origin: languages/js/tests/js/test_js_dataview_get_set_endianness_methods.rs

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

const buffer = new ArrayBuffer(2);
const dv = new DataView(buffer);
dv.setInt16(0, 0x1234, false); // Big-Endian write: 0x12, 0x34

__check(__line(`BigEndian=0x${dv.getInt16(0, false).toString(16)}|LittleEndian=0x${dv.getInt16(0, true).toString(16)}`), "BigEndian=0x1234|LittleEndian=0x3412");
