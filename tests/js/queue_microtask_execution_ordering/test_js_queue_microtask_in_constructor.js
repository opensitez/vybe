// vybe-test: js/queue_microtask_execution_ordering/test_js_queue_microtask_in_constructor
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

class TaskRunner {
    constructor() {
        queueMicrotask(() => console.log("Constructor Microtask"));
    }
}
new TaskRunner();
console.log("Sync Constructor Done");
