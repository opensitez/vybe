// vybe-test: js/misc_es_features/computed_property_in_object_literal
// origin: languages/js/tests/js/test_misc_es_features.rs

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

const prefix = "prop";
const obj = { [prefix + "1"]: "a", [prefix + "2"]: "b" };
__check(__line(obj.prop1, obj.prop2), "a b");
