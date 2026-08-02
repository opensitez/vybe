// vybe-test: js/promise_microtask_order/promise_resolve_in_then_creates_new_promise
// origin: languages/js/tests/js/test_promise_microtask_order.rs

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

Promise.resolve(1).then(()=>Promise.resolve(2)).then(v=>console.log(v));
