// vybe-test: js/ecma_operators/in_operator_with_symbol_in_prototype_chain
// origin: languages/js/tests/js/test_ecma_operators.rs

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

const key = Symbol("token");
const proto = { [key]: "proto" };
const obj = Object.create(proto);
__check(__line(key in obj), "true");
__check(__line(Object.getOwnPropertySymbols(obj).length), "0");
