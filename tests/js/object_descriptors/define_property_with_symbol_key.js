// vybe-test: js/object_descriptors/define_property_with_symbol_key
// origin: languages/js/tests/js/test_object_descriptors.rs

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

const sym = Symbol("id");
const obj = {};
Object.defineProperty(obj, sym, { value: 99, enumerable: true, configurable: true, writable: true });
__check(__line(obj[sym]), "99");
