// vybe-test: js/queue_microtask_execution_ordering/test_js_queue_microtask_inside_async_function
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

async function run() {
    console.log("Async Start");
    queueMicrotask(() => console.log("Queue inside Async"));
    await Promise.resolve();
    console.log("Async After Await");
}
run();
console.log("Main Thread End");
