// vybe-test: js/object_advanced/object_ownpropertysymbols_with_symbol_key
// origin: languages/js/tests/js/test_object_advanced.rs

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

const sym = Symbol("k");
let obj = { a: 1 };
obj[sym] = "secret";
__check(__line(obj[sym]), "secret");
__check(__line(Object.getOwnPropertyNames(obj).includes("Symbol(k)")), "false");
__check(__line(Object.getOwnPropertySymbols(obj).length), "1");
