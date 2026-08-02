// vybe-test: js/object_descriptors/get_own_property_symbols_returns_symbol_keys
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

const s1 = Symbol("a");
const s2 = Symbol("b");
const obj = { [s1]: 1, [s2]: 2, str: 3 };
const syms = Object.getOwnPropertySymbols(obj);
__check(__line(syms.length), "2");
