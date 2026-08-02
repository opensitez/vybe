// vybe-test: js/dataview_arraybuffer/arraybuffer_initial_bytes_are_zero
// origin: languages/js/tests/js/test_dataview_arraybuffer.rs

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

const buf = new ArrayBuffer(4);
const view = new Uint8Array(buf);
__check(__line(view[0], view[1], view[2], view[3]), "0 0 0 0");
