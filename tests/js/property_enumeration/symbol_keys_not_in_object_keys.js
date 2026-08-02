// vybe-test: js/property_enumeration/symbol_keys_not_in_object_keys
// origin: languages/js/tests/js/test_property_enumeration.rs

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

const sym = Symbol("x");
const obj = { [sym]: 1, a: 2, b: 3 };
__check(__line(Object.keys(obj).join(",")), "a,b,Symbol(x)");
__check(__line(Object.getOwnPropertySymbols(obj).length), "1");
