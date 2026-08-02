// vybe-test: js/class_private_in_operator_brand_check/test_js_private_in_operator_inside_instance_method
// origin: languages/js/tests/js/test_js_class_private_in_operator_brand_check.rs

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
    #val;
    constructor(v) { this.#val = v; }
    isSameClass(other) {
        return #val in other;
    }
}
const n1 = new Node(1);
const n2 = new Node(2);
__check(__line(n1.isSameClass(n2) + "|" + n1.isSameClass({})), "true|false");
