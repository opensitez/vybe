// vybe-test: js/symbol_registry_matrix/symbol_keys_are_skipped_by_object_keys
// origin: languages/js/tests/js/test_symbol_registry_matrix.rs

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

const s = Symbol("id");
const obj = { visible: 1, [s]: 2 };
__check(__line(Object.keys(obj).join(",")), "visible");
