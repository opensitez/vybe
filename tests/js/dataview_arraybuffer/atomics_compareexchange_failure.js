// vybe-test: js/dataview_arraybuffer/atomics_compareexchange_failure
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

const sab = new SharedArrayBuffer(4);
const ia = new Int32Array(sab);
ia[0] = 5;
const old = Atomics.compareExchange(ia, 0, 99, 10);
__check(__line(old), "5");
__check(__line(ia[0]), "5");
