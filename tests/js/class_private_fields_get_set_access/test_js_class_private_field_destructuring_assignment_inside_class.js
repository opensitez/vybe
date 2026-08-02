// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_destructuring_assignment_inside_class
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

class DataHolder {
    #val = 100;
    swap(other) {
        [this.#val, other.#val] = [other.#val, this.#val];
    }
    getVal() { return this.#val; }
}
const d1 = new DataHolder();
const d2 = new DataHolder();
d1.swap(d2);
__check(__line(`${d1.getVal()}:${d2.getVal()}`), "100:100");
