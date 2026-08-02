// vybe-test: js/property_enumeration/reflect_ownkeys_order_with_integers_and_symbols
// origin: languages/js/tests/js/test_property_enumeration.rs

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

const s1 = Symbol("first");
const s2 = Symbol("second");
const obj = { "10": "ten", "a": "ay", "2": "two", [s1]: "one", [s2]: "two" };
const keys = Reflect.ownKeys(obj);
__check(__line(keys.length), "5");
__check(__line(keys[0]), "2");
__check(__line(keys[1]), "10");
__check(__line(keys[2]), "a");
__check(__line(typeof keys[3]), "symbol");
