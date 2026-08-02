// vybe-test: js/property_ordering/reflect_own_keys_all_types_ordered
// origin: languages/js/tests/js/test_property_ordering.rs

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
const obj = { 1: "b", sym: "s", 0: "a" };
obj[sym] = "sym";
const names = Object.getOwnPropertyNames(obj);
const intKeys = names.filter(k => /^\d+$/.test(k)).sort((a,b) => +a - +b);
const strKeys = names.filter(k => !/^\d+$/.test(k));
const symKeys = Object.getOwnPropertySymbols(obj);
__check(__line(intKeys[0]), "0"); // "0"
__check(__line(intKeys[1]), "1"); // "1"
__check(__line(strKeys[0]), "sym"); // "sym"
__check(__line(typeof symKeys[0]), "symbol"); // "symbol"
