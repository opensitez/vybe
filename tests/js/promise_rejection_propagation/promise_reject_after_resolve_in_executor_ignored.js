// vybe-test: js/promise_rejection_propagation/promise_reject_after_resolve_in_executor_ignored
// origin: languages/js/tests/js/test_promise_rejection_propagation.rs

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

new Promise((res,rej)=>{res(1);rej(2);}).then(v=>console.log(v));
