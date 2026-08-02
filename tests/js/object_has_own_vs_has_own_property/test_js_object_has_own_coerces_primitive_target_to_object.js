// vybe-test: js/object_has_own_vs_has_own_property/test_js_object_has_own_coerces_primitive_target_to_object
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

__check(__line(`${Object.hasOwn("hello", 0)}:${Object.hasOwn("hello", "length")}:${Object.hasOwn(123, "toString")}`), "true:true:false");
