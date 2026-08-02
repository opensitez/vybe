// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_initializer_expression_evaluation
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

let idCounter = 0;
class Item {
    #id = ++idCounter;
    getId() { return this.#id; }
}
const i1 = new Item();
const i2 = new Item();
__check(__line(`${i1.getId()}:${i2.getId()}`), "1:2");
