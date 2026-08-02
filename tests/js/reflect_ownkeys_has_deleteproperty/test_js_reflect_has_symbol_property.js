// vybe-test: js/reflect_ownkeys_has_deleteproperty/test_js_reflect_has_symbol_property
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

const sym = Symbol("sym");
const obj = { [sym]: "Val" };
__check(__line(Reflect.has(obj, sym)), "true");
