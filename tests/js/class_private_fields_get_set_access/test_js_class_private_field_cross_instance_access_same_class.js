// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_cross_instance_access_same_class
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

class Vector {
    #x; #y;
    constructor(x, y) { this.#x = x; this.#y = y; }
    add(otherVector) {
        return new Vector(this.#x + otherVector.#x, this.#y + otherVector.#y);
    }
    toString() { return `(${this.#x},${this.#y})`; }
}
const v1 = new Vector(1, 2);
const v2 = new Vector(3, 4);
__check(__line(v1.add(v2).toString()), "(4,6)");
