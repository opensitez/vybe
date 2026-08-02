// vybe-test: js/structured_clone_map_set_regexp_date/test_js_structured_clone_map_key_value_identity_preservation
// origin: languages/js/tests/js/test_js_structured_clone_map_set_regexp_date.rs

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

const shared = { id: 1 };
const map = new Map([[shared, shared]]);
const clone = structuredClone(map);
const [cloneKey, cloneVal] = [...clone.entries()][0];
__check(__line((cloneKey !== shared) + "|" + (cloneKey === cloneVal)), "true|true");
