// vybe-test: js/atomics_compare_exchange_store_load/test_js_atomics_compare_exchange_bigint64
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

const sab = new SharedArrayBuffer(8);
const bi64 = new BigInt64Array(sab);
bi64[0] = 100n;
const old = Atomics.compareExchange(bi64, 0, 100n, 500n);
__check(__line(old.toString() + "|" + bi64[0].toString()), "100|500");
