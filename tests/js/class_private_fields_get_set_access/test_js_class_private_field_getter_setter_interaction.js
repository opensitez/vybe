// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_getter_setter_interaction
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

class Range {
    #val = 0;
    set value(v) {
        if (v >= 0) this.#val = v;
    }
    get value() { return this.#val; }
}
const r = new Range();
r.value = 50;
r.value = -10; // Rejected by setter
__check(__line(r.value), "50");
