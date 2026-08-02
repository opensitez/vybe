// vybe-test: js/dataview_arraybuffer/dataview_bytelength_and_byteoffset
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

const buf = new ArrayBuffer(16);
const dv = new DataView(buf, 4, 8);
__check(__line(dv.byteLength), "8");
__check(__line(dv.byteOffset), "4");
