// vybe-test: js/weakref_weakmap_advanced/finalizationregistry_constructor_exists
// origin: languages/js/tests/js/test_weakref_weakmap_advanced.rs

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

const registry = new FinalizationRegistry(val => {});
__check(__line(typeof registry.register), "function");
__check(__line(typeof registry.unregister), "function");
