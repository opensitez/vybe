// vybe-test: js/promise_microtask_order/catch_on_already_rejected_runs_microtask
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

const p=Promise.reject(1); const o=[]; p.catch(()=>o.push("c")); o.push("s"); console.log(o.join(","));
