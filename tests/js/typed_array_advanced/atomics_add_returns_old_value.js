// vybe-test: js/typed_array_advanced/atomics_add_returns_old_value
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

const sab = new SharedArrayBuffer(4);
const ta = new Int32Array(sab);
ta[0] = 10;
const old = Atomics.add(ta, 0, 5);
__check(__line(old), "10");      // old value
__check(__line(ta[0]), "15");    // new value
