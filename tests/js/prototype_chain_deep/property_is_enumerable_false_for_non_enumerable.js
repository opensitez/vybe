// vybe-test: js/prototype_chain_deep/property_is_enumerable_false_for_non_enumerable
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

const obj = {};
Object.defineProperty(obj, "x", { value: 1, enumerable: false, configurable: true });
__check(__line(obj.propertyIsEnumerable("x")), "false");
