// vybe-test: js/prototype_chain_deep/object_keys_vs_own_property_names_differ_for_non_enumerable
// origin: languages/js/tests/js/test_prototype_chain_deep.rs

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

const obj = { a: 1 };
Object.defineProperty(obj, "b", { value: 2, enumerable: false, configurable: true });
__check(__line(Object.keys(obj).join(",")), "a");
__check(__line(Object.getOwnPropertyNames(obj).sort().join(",")), "a,b");
