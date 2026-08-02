// vybe-test: js/queue_microtask_execution_ordering/test_js_queue_microtask_fifo_ordering
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

for (let i = 1; i <= 3; i++) {
    queueMicrotask(() => console.log("Task " + i));
}
