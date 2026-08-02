// vybe-test: js/symbol_protocols/symbol_nonEnumerable_hiding
// origin: languages/js/tests/js/test_symbol_protocols.rs

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

const SECRET = Symbol("secret");
const obj = {
    name: "Alice",
    [SECRET]: "hidden",
    age: 30
};
__check(__line(Object.keys(obj).join(",")), "name,age");
__check(__line(obj[SECRET]), "hidden");
__check(__line(Object.getOwnPropertySymbols(obj).length), "1");
