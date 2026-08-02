// vybe-test: js/object_prototype_has_own_property_is_prototype_of/test_js_object_property_is_enumerable_symbol_keys
// origin: languages/js/tests/js/test_js_object_prototype_has_own_property_is_prototype_of.rs

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

const s1 = Symbol("s1");
const s2 = Symbol("s2");
const obj = { [s1]: 10 };
Object.defineProperty(obj, s2, { value: 20, enumerable: false });

__check(__line(obj.propertyIsEnumerable(s1)), "true");
__check(__line(obj.propertyIsEnumerable(s2)), "false");
