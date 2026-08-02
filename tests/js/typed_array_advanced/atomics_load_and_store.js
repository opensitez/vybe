// vybe-test: js/typed_array_advanced/atomics_load_and_store
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
Atomics.store(ta, 0, 777);
__check(__line(Atomics.load(ta, 0)), "777");
