// vybe-test: js/structured_clone_map_set_regexp_date/test_js_structured_clone_boolean_object_wrapper
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

const boolObj = new Boolean(true);
const clone = structuredClone(boolObj);
__check(__line((clone instanceof Boolean) + "|" + (clone.valueOf() === true)), "true|true");
