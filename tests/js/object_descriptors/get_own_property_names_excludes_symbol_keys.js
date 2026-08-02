// vybe-test: js/object_descriptors/get_own_property_names_excludes_symbol_keys
// origin: languages/js/tests/js/test_object_descriptors.rs

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
const obj = { a: 1, [sym]: 2 };
const names = Object.getOwnPropertyNames(obj);
__check(__line(names.includes("a")), "true");
__check(__line(names.length), "1");
