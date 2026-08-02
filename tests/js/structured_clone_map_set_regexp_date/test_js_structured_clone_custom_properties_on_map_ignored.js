// vybe-test: js/structured_clone_map_set_regexp_date/test_js_structured_clone_custom_properties_on_map_ignored
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

const map = new Map();
map.customProp = "customData";
const clone = structuredClone(map);
__check(__line(clone.size + "|hasCustomProp=" + Object.hasOwn(clone, "customProp")), "0|hasCustomProp=false");
