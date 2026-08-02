// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_spin_lock_simulation
// origin: languages/js/tests/js/test_js_atomics_compare_exchange_store_load.rs

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

const lock = new Int32Array(new SharedArrayBuffer(4));
// Try acquire lock (0 -> 1)
const acquired = Atomics.compareExchange(lock, 0, 0, 1) === 0;
// Release lock (1 -> 0)
if (acquired) Atomics.store(lock, 0, 0);
__check(__line(acquired + "|" + Atomics.load(lock, 0)), "true|0");
