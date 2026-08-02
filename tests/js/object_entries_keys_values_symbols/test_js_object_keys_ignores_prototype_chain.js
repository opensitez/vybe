// vybe-test: js/object_entries_keys_values_symbols/test_js_object_keys_ignores_prototype_chain
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

const proto = { protoKey: 1 };
const obj = Object.create(proto);
obj.ownKey = 2;
__check(__line(Object.keys(obj).join(",")), "ownKey");
