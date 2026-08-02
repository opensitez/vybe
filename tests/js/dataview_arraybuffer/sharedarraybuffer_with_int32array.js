// vybe-test: js/dataview_arraybuffer/sharedarraybuffer_with_int32array
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

const sab = new SharedArrayBuffer(16);
const ia = new Int32Array(sab);
ia[0] = 42;
__check(__line(ia[0]), "42");
