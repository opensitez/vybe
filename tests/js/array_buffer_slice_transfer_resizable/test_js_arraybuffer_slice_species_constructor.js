// vybe-test: js/array_buffer_slice_transfer_resizable/test_js_arraybuffer_slice_species_constructor
// origin: languages/js/tests/js/test_js_array_buffer_slice_transfer_resizable.rs

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

class CustomBuffer extends ArrayBuffer {}
const buf = new CustomBuffer(16);
const sliced = buf.slice(0, 8);
__check(__line(sliced.byteLength + "|isCustom=" + (sliced instanceof CustomBuffer)), "8|isCustom=true");
