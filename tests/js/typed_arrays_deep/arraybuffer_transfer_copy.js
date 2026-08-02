// vybe-test: js/typed_arrays_deep/arraybuffer_transfer_copy
// origin: languages/js/tests/js/test_typed_arrays_deep.rs

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

const buf1 = new ArrayBuffer(4);
const view1 = new Uint32Array(buf1);
view1[0] = 42;
// Copy via typed array
const buf2 = buf1.slice(0);
const view2 = new Uint32Array(buf2);
view2[0] = 99;
__check(__line(view1[0]), "42");
__check(__line(view2[0]), "99");
