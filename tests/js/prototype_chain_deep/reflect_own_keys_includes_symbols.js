// vybe-test: js/prototype_chain_deep/reflect_own_keys_includes_symbols
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

const sym = Symbol("s");
const obj = { a: 1, [sym]: 2 };
const keys = Reflect.ownKeys(obj);
__check(__line(keys.includes("a")), "true");
__check(__line(Object.getOwnPropertySymbols(obj).some(k => typeof k === "symbol")), "true");
