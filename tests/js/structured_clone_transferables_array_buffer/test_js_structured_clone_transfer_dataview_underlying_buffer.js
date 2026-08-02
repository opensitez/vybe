// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_dataview_underlying_buffer
// origin: languages/js/tests/js/test_js_structured_clone_transferables_array_buffer.rs

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
const dv = new DataView(buf);
dv.setInt32(0, 9999);
const cloneDV = structuredClone(dv, { transfer: [buf] });

__check(__line((buf.byteLength === 0) + "|" + cloneDV.getInt32(0)), "true|9999");
