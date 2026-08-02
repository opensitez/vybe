// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_preserves_nested_cyclical_identity
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

const buf = new Uint8Array([7, 8, 9]).buffer;
const node = { buffer: buf };
node.self = node;
const clone = structuredClone(node, { transfer: [buf] });

__check(__line((buf.byteLength === 0) + "|" + (clone.self === clone) + "|" + new Uint8Array(clone.buffer).join(",")), "true|true|7,8,9");
