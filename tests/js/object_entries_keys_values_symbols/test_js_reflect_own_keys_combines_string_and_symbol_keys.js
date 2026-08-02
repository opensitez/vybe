// vybe-test: js/object_entries_keys_values_symbols/test_js_reflect_own_keys_combines_string_and_symbol_keys
// origin: languages/js/tests/js/test_js_object_entries_keys_values_symbols.rs

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

const s = Symbol("sym");
const obj = { b: 2, 1: "num", [s]: "symVal" };
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
const keys = Reflect.ownKeys(obj);
__check(__line(keys.map(String).join(",")), "1,b,hidden,Symbol(sym)");
