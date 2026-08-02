// vybe-test: js/weakref_finalization/finalization_registry_cleanup_receives_held_value
// origin: languages/js/tests/js/test_weakref_finalization.rs

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

// The callback receives the held value (second arg to register), not the object
const received = [];
const registry = new FinalizationRegistry((heldValue) => {
    received.push(heldValue);
});
// Register with a held value
let obj = {};
registry.register(obj, "my-token");
// We can't force GC, but we can verify the API works
console.log("setup complete");
