// vybe-test: js/shared_array_buffer_view_sharing/test_js_shared_array_buffer_byte_length_property_descriptor
// origin: languages/js/tests/js/test_js_shared_array_buffer_view_sharing.rs

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

const desc = Object.getOwnPropertyDescriptor(SharedArrayBuffer.prototype, "byteLength");
__check(__line(`${typeof desc.get}:${desc.set}:${desc.enumerable}:${desc.configurable}`), "function:undefined:false:true");
