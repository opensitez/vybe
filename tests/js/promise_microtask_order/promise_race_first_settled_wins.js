// vybe-test: js/promise_microtask_order/promise_race_first_settled_wins
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

Promise.race([new Promise(r=>setTimeout(()=>r("slow"),10)),Promise.resolve("fast")]).then(v=>console.log(v));
