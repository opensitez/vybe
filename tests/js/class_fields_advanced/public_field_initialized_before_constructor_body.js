// vybe-test: js/class_fields_advanced/public_field_initialized_before_constructor_body
// origin: languages/js/tests/js/test_class_fields_advanced.rs

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

class Child {
    compute() { return 5; }
    constructor() { this.value = this.compute(); }
}
const c = new Child();
__check(__line(c.value), "5");
