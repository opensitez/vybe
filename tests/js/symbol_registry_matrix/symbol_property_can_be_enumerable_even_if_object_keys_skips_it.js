// vybe-test: js/symbol_registry_matrix/symbol_property_can_be_enumerable_even_if_object_keys_skips_it
// origin: languages/js/tests/js/test_symbol_registry_matrix.rs

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

const s = Symbol("a");
const obj = {};
obj[s] = 1;
__check(__line(obj.propertyIsEnumerable(s)), "true");
