// vybe-test: js/typed_array_advanced/arraybuffer_slice_creates_copy
// origin: languages/js/tests/js/test_typed_array_advanced.rs

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
const copy = buf.slice(0, 4);
const copyView = new Uint8Array(copy);
copyView[0] = 99;
__check(__line(view[0]), "42"); // original unchanged
__check(__line(copyView[0]), "99");
