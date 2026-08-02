// vybe-test: js/atomics_wait_notify_async_wait/test_js_atomics_wait_async_promise_resolution
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

if (typeof Atomics.waitAsync === "function") {
    const sab = new SharedArrayBuffer(4);
    const i32 = new Int32Array(sab);
    i32[0] = 1;
    const res = Atomics.waitAsync(i32, 0, 1, 1);
    if (res.async) {
        (async () => {
            console.log(await res.value);
        })();
    } else {
        console.log(res.value);
    }
} else {
    console.log("timed-out");
}
