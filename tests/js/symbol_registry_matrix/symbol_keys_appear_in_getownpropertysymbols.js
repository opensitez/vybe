// vybe-test: js/symbol_registry_matrix/symbol_keys_appear_in_getownpropertysymbols
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

const a = Symbol("a");
const b = Symbol("b");
const obj = { [a]: 1, [b]: 2 };
__check(__line(Object.getOwnPropertySymbols(obj).length), "2");
