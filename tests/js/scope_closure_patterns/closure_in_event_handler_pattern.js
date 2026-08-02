// vybe-test: js/scope_closure_patterns/closure_in_event_handler_pattern
// origin: languages/js/tests/js/test_scope_closure_patterns.rs

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

function attachHandlers(items) {
    return items.map((item, index) => ({
        name: item,
        handler: () => `Clicked ${item} at index ${index}`
    }));
}
const handlers = attachHandlers(["a", "b", "c"]);
__check(__line(handlers[0].handler()), "Clicked a at index 0");
__check(__line(handlers[2].handler()), "Clicked c at index 2");
