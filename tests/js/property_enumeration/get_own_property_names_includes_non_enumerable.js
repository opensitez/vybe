// vybe-test: js/property_enumeration/get_own_property_names_includes_non_enumerable
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

const obj = { a: 1 };
Object.defineProperty(obj, "b", { value: 2, enumerable: false });
const all = Object.getOwnPropertyNames(obj).sort();
const enumOnly = Object.keys(obj);
__check(__line(all.join(",")), "a,b");
__check(__line(enumOnly.join(",")), "a");
