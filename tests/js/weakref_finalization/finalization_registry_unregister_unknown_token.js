// vybe-test: js/weakref_finalization/finalization_registry_unregister_unknown_token
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

const registry = new FinalizationRegistry(() => {});
const token = {};
const result = registry.unregister(token); // was never registered
__check(__line(result), "false");
