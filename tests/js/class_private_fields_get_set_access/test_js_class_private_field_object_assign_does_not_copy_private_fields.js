// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_object_assign_does_not_copy_private_fields
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

class Container {
    #secret = 999;
    publicVal = 100;
    getSecret() { return this.#secret; }
}
const c1 = new Container();
const c2 = Object.assign({}, c1);
__check(__line(c2.publicVal + "|" + (typeof c2.getSecret === "undefined")), "100|true");
