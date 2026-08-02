// vybe-test: js/structured_clone_map_set_regexp_date/test_js_structured_clone_set_with_nested_maps
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

const innerMap = new Map([["x", 100]]);
const set = new Set([innerMap]);
const clone = structuredClone(set);
const clonedMap = [...clone][0];
__check(__line(clonedMap.get("x")), "100");
