// vybe-test: js/object_descriptors/object_assign_symbol_properties_are_copied
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

const s = Symbol("s");
const source = { a: 1 };
source[s] = 2;
const result = Object.assign({}, source, { b: 3 });
__check(__line(result.a), "1");
__check(__line(result.b), "3");
__check(__line(result[s]), "2");
__check(__line(Object.getOwnPropertySymbols(result).length), "1");
