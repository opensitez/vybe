// vybe-test: js/object_has_own_vs_has_own_property/test_js_object_has_own_coerces_property_key_to_string_or_symbol
// origin: languages/js/tests/js/test_js_object_has_own_vs_has_own_property.rs

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

const obj = { 42: "answer" };
__check(__line(Object.hasOwn(obj, 42) + "|" + Object.hasOwn(obj, "42")), "true|true");
