// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_iterable_options
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

const buf = new Uint8Array([42]).buffer;
const transferSet = new Set([buf]);
const clone = structuredClone(buf, { transfer: transferSet });
__check(__line((buf.byteLength === 0) + "|" + new Uint8Array(clone)[0]), "true|42");
