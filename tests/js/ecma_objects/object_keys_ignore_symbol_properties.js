// vybe-test: js/ecma_objects/object_keys_ignore_symbol_properties
// origin: languages/js/tests/js/test_ecma_objects.rs

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

const id = Symbol("id");
const obj = { a: 1, [id]: 2 };
__check(__line(Object.keys(obj).join(",")), "a");
__check(__line(obj[id]), "2");
