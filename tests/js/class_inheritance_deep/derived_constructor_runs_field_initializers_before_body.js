// vybe-test: js/class_inheritance_deep/derived_constructor_runs_field_initializers_before_body
// origin: languages/js/tests/js/test_class_inheritance_deep.rs

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

const order = [];
class Base {
    constructor() { order.push("base"); }
}
class Derived extends Base {
    initialized = (order.push("field"), 1);
    constructor() {
        super();
        order.push("constructor");
    }
}
new Derived();
__check(__line(order.join("|")), "base|field|constructor");
