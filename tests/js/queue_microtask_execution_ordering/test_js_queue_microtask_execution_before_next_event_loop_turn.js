// vybe-test: js/queue_microtask_execution_ordering/test_js_queue_microtask_execution_before_next_event_loop_turn
// origin: languages/js/tests/js/test_js_queue_microtask_execution_ordering.rs

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

let order = [];
queueMicrotask(() => order.push("Micro1"));
queueMicrotask(() => order.push("Micro2"));
Promise.resolve().then(() => {
    console.log(order.join("->"));
});
