// vybe-test: js/structured_clone_circular_references/test_js_structured_clone_symbol_property_ignored_or_error
// origin: languages/js/tests/js/test_js_structured_clone_circular_references.rs

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

const sym = Symbol("key");
const obj = { [sym]: "data", stringKey: "data" };
const clone = structuredClone(obj);
__check(__line(clone.stringKey + "|hasSym=" + (sym in clone)), "data|hasSym=false");
