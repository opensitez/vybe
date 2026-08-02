// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_not_enumerable_in_keys
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

class User {
    #id = 1;
    name = "Alice";
}
const u = new User();
__check(__line(Object.keys(u).join(",") + "|Count=" + Object.getOwnPropertyNames(u).length), "name|Count=1");
