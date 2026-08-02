// vybe-test: js/prototype_chain_shadowing_property_lookup/test_js_prototype_chain_symbol_property_shadowing_and_reflection
// origin: languages/js/tests/js/test_js_prototype_chain_shadowing_property_lookup.rs

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

const token = Symbol("token");
const proto = { [token]: "proto" };
const obj = Object.create(proto);

obj[token] = "own";
__check(__line(Object.getOwnPropertySymbols(obj).length), "1");
__check(__line(Object.getOwnPropertySymbols(obj)[0] === token), "true");
__check(__line(token in obj), "true");
__check(__line(obj[token]), "own");
