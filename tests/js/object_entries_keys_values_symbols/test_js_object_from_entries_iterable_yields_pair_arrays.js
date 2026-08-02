// vybe-test: js/object_entries_keys_values_symbols/test_js_object_from_entries_iterable_yields_pair_arrays
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

function* pairGen() {
    yield ["x", 1];
    yield ["y", 2];
}
const obj = Object.fromEntries(pairGen());
__check(__line(`${obj.x}:${obj.y}`), "1:2");
