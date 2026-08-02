// vybe-test: js/symbol_for_key_for_registry/test_js_symbol_keys_ignored_by_for_in_and_object_keys
// origin: languages/js/tests/js/test_js_symbol_for_key_for_registry.rs

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

const sym = Symbol("hidden");
const obj = { [sym]: 10, pub: 20 };
__check(__line(Object.keys(obj).join(",") + "|hasIn=" + (sym in obj)), "pub|hasIn=true");
