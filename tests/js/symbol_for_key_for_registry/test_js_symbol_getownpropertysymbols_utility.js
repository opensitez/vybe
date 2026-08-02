// vybe-test: js/symbol_for_key_for_registry/test_js_symbol_getownpropertysymbols_utility
// origin: languages/js/tests/js/test_js_symbol_for_key_for_registry.rs

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
const s2 = Symbol.for("b");
const obj = { [s1]: 1, [s2]: 2, stringKey: 3 };

const symbols = Object.getOwnPropertySymbols(obj);
console.log(symbols.length + "|" + (symbols[0] === s1) + "|" + (symbols[1] === s2));
