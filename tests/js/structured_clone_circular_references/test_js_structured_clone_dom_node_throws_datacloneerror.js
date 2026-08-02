// vybe-test: js/structured_clone_circular_references/test_js_structured_clone_dom_node_throws_datacloneerror
// origin: languages/js/tests/js/test_js_structured_clone_circular_references.rs

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

try {
    structuredClone({ symbolVal: Symbol() }); // Symbol values cannot be cloned!
} catch (e) {
    __check(__line("DataCloneError Symbol"), "DataCloneError Symbol");
}
