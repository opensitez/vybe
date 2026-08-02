// vybe-test: js/symbol_registry_matrix/object_getownpropertysymbols_preserves_insertion_order
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
const obj = {};
obj[a] = 1;
obj[b] = 2;
const syms = Object.getOwnPropertySymbols(obj);
__check(__line(syms[0] === a), "true");
__check(__line(syms[1] === b), "true");
