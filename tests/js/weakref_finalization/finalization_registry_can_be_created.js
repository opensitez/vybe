// vybe-test: js/weakref_finalization/finalization_registry_can_be_created
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

const registry = new FinalizationRegistry((value) => {
    console.log("cleaned:" + value);
});
console.log(registry instanceof FinalizationRegistry);
