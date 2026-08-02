// vybe-test: js/atomics_is_lock_free_buffer_size/test_js_atomics_is_lock_free_typed_array_element_bytes
// origin: languages/js/tests/js/test_js_atomics_is_lock_free_buffer_size.rs

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

__check(__line([
    Atomics.isLockFree(Int8Array.BYTES_PER_ELEMENT),
    Atomics.isLockFree(Int16Array.BYTES_PER_ELEMENT),
    Atomics.isLockFree(Int32Array.BYTES_PER_ELEMENT),
    Atomics.isLockFree(BigInt64Array.BYTES_PER_ELEMENT)
].every(res => typeof res === "boolean")), "true");
