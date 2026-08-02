// vybe-test: js/prototype_chain_advanced/symbol_property_in_prototype_chain_can_be_shadowed_by_own_symbol_property
// origin: languages/js/tests/js/test_prototype_chain_advanced.rs

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
const proto = {
    [token]: "proto",
    toString() { return "from proto"; }
};
const obj = Object.create(proto);
obj[token] = "own";
__check(__line(Object.hasOwn(obj, token)), "true");
__check(__line(obj[token]), "own");
