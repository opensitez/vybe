// vybe-test: js/weakref_weakmap_advanced/finalizationregistry_register_object
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

let collected = false;
const registry = new FinalizationRegistry(val => { collected = true; });
let obj = { data: "hello" };
registry.register(obj, "cleanup token");
__check(__line(typeof obj.data), "string");
