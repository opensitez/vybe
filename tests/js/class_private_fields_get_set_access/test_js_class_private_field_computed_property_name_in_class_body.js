// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_computed_property_name_in_class_body
// origin: languages/js/tests/js/test_js_class_private_fields_get_set_access.rs

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

const propName = "dynamic";
class Component {
    #data = 500;
    [propName]() { return this.#data; }
}
__check(__line(new Component().dynamic()), "500");
