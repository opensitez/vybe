// vybe-test: js/property_enumeration/reflect_ownkeys_includes_symbols_and_all_strings
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

const sym = Symbol("s");
const obj = {};
Object.defineProperty(obj, "hidden", { value: 1, enumerable: false });
obj.visible = 2;
obj[sym] = 3;
const all = Reflect.ownKeys(obj);
__check(__line(all.includes("hidden")), "true");
__check(__line(all.includes("visible")), "true");
__check(__line(Object.getOwnPropertySymbols(obj).length > 0), "true");
