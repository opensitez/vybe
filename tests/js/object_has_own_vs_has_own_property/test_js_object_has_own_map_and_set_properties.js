// vybe-test: js/object_has_own_vs_has_own_property/test_js_object_has_own_map_and_set_properties
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

const set = new Set();
__check(__line(`${Object.hasOwn(set, "size")}:${Object.hasOwn(Set.prototype, "size")}`), "false:true");
