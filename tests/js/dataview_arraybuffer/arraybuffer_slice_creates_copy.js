// vybe-test: js/dataview_arraybuffer/arraybuffer_slice_creates_copy
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

const buf = new ArrayBuffer(8);
const view = new Uint8Array(buf);
view[0] = 42;
const sliced = buf.slice(0, 4);
const slicedView = new Uint8Array(sliced);
slicedView[0] = 99;
__check(__line(view[0]), "42");
__check(__line(slicedView[0]), "99");
