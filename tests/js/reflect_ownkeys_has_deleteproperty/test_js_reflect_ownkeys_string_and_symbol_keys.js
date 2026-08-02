// vybe-test: js/reflect_ownkeys_has_deleteproperty/test_js_reflect_ownkeys_string_and_symbol_keys
// origin: languages/js/tests/js/test_js_reflect_ownkeys_has_deleteproperty.rs

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
const obj = { a: 1, 0: "zero", [sym]: 2 };
const keys = Reflect.ownKeys(obj);
__check(__line(keys.map(k => String(k)).join(",")), "0,a,Symbol(s)");
