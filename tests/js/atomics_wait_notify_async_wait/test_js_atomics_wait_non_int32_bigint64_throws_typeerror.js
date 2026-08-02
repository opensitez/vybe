// vybe-test: js/atomics_wait_notify_async_wait/test_js_atomics_wait_non_int32_bigint64_throws_typeerror
// origin: languages/js/tests/js/test_js_atomics_wait_notify_async_wait.rs

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

const sab = new SharedArrayBuffer(2);
const i16 = new Int16Array(sab);
try {
    Atomics.wait(i16, 0, 0); // Atomics.wait requires Int32Array or BigInt64Array!
} catch (e) {
    __check(__line("Atomics.wait Invalid TypedArray TypeError"), "Atomics.wait Invalid TypedArray TypeError");
}
