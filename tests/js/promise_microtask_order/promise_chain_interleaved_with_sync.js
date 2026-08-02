// vybe-test: js/promise_microtask_order/promise_chain_interleaved_with_sync
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

const o=[]; Promise.resolve().then(()=>o.push(1)).then(()=>o.push(2)); Promise.resolve().then(()=>o.push(3)); o.push(0); console.log(o.join(","));
