// vybe-test: js/atomics_wait_notify_async_wait/test_js_atomics_notify_property_descriptor
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

const desc = Object.getOwnPropertyDescriptor(Atomics, "notify");
__check(__line(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${Atomics.notify.length}`), "true:false:true:3");
