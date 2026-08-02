// vybe-test: js/promise_microtask_order/nested_promise_then_order_fifo
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

const o=[]; Promise.resolve().then(()=>o.push("a")).then(()=>o.push("b")); o.push("s"); console.log(o.join(","));
