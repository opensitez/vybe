// vybe-test: js/queue_microtask_execution_ordering/test_js_queue_microtask_clearing_local_references
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

let temp = { heavyData: [1, 2, 3] };
queueMicrotask(() => {
    temp = null;
    console.log("Cleaned Reference: " + (temp === null));
});
