// vybe-test: js/class_static_private_fields_methods/test_js_class_static_private_field_not_in_keys_or_get_own_property_names
// origin: languages/js/tests/js/test_js_class_static_private_fields_methods.rs

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

class Sample {
    static #priv = 1;
    static pub = 2;
}
__check(__line(Object.keys(Sample).join(",") + "|Count=" + Object.getOwnPropertyNames(Sample).length), "pub|Count=4");
