// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_multiple_declarations
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

class Node {
    #left; #right; #val;
    constructor(v) { this.#val = v; }
    setChildren(l, r) { this.#left = l; this.#right = r; }
    getSummary() {
        return `${this.#val}:[${this.#left.#val},${this.#right.#val}]`;
    }
}
const parent = new Node("Root");
parent.setChildren(new Node("L"), new Node("R"));
__check(__line(parent.getSummary()), "Root:[L,R]");
